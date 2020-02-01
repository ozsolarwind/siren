#!/usr/bin/python3
#
#  Copyright (C) 2018-2020 Sustainable Energy Now Inc., Angus King
#
#  powermatch.py - This file is possibly part of SIREN.
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
from PyQt4 import QtCore, QtGui
import displayobject
import displaytable
from credits import fileVersion
from matplotlib import rcParams
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
import xlrd
import configparser  # decode .ini file
from zoompan import ZoomPanX

tech_names = ['Load', 'Onshore Wind', 'Offshore Wind', 'Rooftop PV', 'Fixed PV', 'Single Axis PV',
              'Dual Axis PV', 'Biomass', 'Geothermal', 'Other1', 'CST', 'Shortfall']
col_letters = ' ABCDEFGHIJKLMNOPQRSTUVWXYZ'
# same order as self.file_labels
C = 0 # Constraints
G = 1 # Generators
O = 2 # Optimisation
D = 3 # Data
R = 4 # Results
def ss_col(col, base=1):
    if base == 1:
        col -= 1
    c1 = col // 26
    c2 = col % 26
    return (col_letters[c1] + col_letters[c2 + 1]).strip()


class ThumbListWidget(QtGui.QListWidget):
    def __init__(self, type, parent=None):
        super(ThumbListWidget, self).__init__(parent)
        self.setIconSize(QtCore.QSize(124, 124))
        self.setDragDropMode(QtGui.QAbstractItemView.DragDrop)
        self.setSelectionMode(QtGui.QAbstractItemView.ExtendedSelection)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            super(ThumbListWidget, self).dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(QtCore.Qt.CopyAction)
            event.accept()
        else:
            super(ThumbListWidget, self).dragMoveEvent(event)

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(QtCore.Qt.CopyAction)
            event.accept()
            links = []
            for url in event.mimeData().urls():
                links.append(str(url.toLocalFile()))
            self.emit(QtCore.SIGNAL("dropped"), links)
        else:
            event.setDropAction(QtCore.Qt.MoveAction)
            super(ThumbListWidget, self).dropEvent(event)


class ClickableQLabel(QtGui.QLabel):
    def __init(self, parent):
        QLabel.__init__(self, parent)

    def mousePressEvent(self, event):
        QtGui.QApplication.widgetAt(event.globalPos()).setFocus()
        self.emit(QtCore.SIGNAL('clicked()'))


class Constraint:
    def __init__(self, name, category, capacity_min, capacity_max, rampup_max, rampdown_max,
                 recharge_max, recharge_loss, discharge_max, discharge_loss, parasitic_loss):
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
            caps = values.split(' ')
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
            caps = values.split(' ')
            try:
                self.capacity_min = float(caps[0])
            except:
                self.capacity_min = 0
            try:
                self.capacity_max = float(caps[1])
            except:
                self.capacity_max = None
            try:
                self.capacity_step = float(caps[2])
            except:
                self.capacity_step = None
            self.capacities = None
        else:
            self.capacity_min = 0
            self.capacity_max = None
            self.capacity_step = None
            self.capacities = None
        self.capacity = None


class Adjustments(QtGui.QDialog):
    def __init__(self, data, adjustin):
        super(Adjustments, self).__init__()
        self.adjusts = {}
        self.checkbox = {}
        self.labels = {}
        self.results = None
        self.grid = QtGui.QGridLayout()
        self.data = {}
        ctr = 0
        for key, capacity in data:
            if key != 'Load' and capacity is None:
                continue
            self.adjusts[key] = QtGui.QDoubleSpinBox()
            self.adjusts[key].setRange(0, 25)
            self.adjusts[key].setDecimals(3)
            try:
                self.adjusts[key].setValue(adjustin[key])
            except:
                self.adjusts[key].setValue(1.)
            self.data[key] = capacity
            self.adjusts[key].setSingleStep(.1)
            self.adjusts[key].setObjectName(key)
            self.grid.addWidget(QtGui.QLabel(key), ctr, 0)
            self.grid.addWidget(self.adjusts[key], ctr, 1)
            if key != 'Load' or key == 'Load':
                self.adjusts[key].valueChanged.connect(self.adjust)
                self.labels[key] = QtGui.QLabel('')
                self.labels[key].setObjectName(key + 'label')
                if key != 'Load':
                    mw = '{:.0f} MW'.format(capacity * self.adjusts[key].value())
                    self.labels[key].setText(mw)
                self.grid.addWidget(self.labels[key], ctr, 2)
            ctr += 1
        quit = QtGui.QPushButton('Quit', self)
        self.grid.addWidget(quit, ctr, 0)
        quit.clicked.connect(self.quitClicked)
        show = QtGui.QPushButton('Proceed', self)
        self.grid.addWidget(show, ctr, 1)
        show.clicked.connect(self.showClicked)
        frame = QtGui.QFrame()
        frame.setLayout(self.grid)
        self.scroll = QtGui.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtGui.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
        self.setWindowTitle('SIREN - Powermatch - Adjust generators')
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        QtGui.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
        self.show()

    def adjust(self):
        key = str(self.sender().objectName())
        mw = '{:.0f} MW'.format(self.data[key] * self.sender().value())
        self.labels[key].setText(mw)

    def closeEvent(self, event):
        event.accept()

    def quitClicked(self):
        self.close()

    def showClicked(self):
        self.results = {}
        for key in list(self.adjusts.keys()):
            self.results[key] = round(self.adjusts[key].value(), 3)
        self.close()

    def getValues(self):
        return self.results

class powerMatch(QtGui.QWidget):

    def get_filename(self, filename):
        if str(filename).find('/') == 0: # full directory in non-Windows
            return filename
        elif (sys.platform == 'win32' or sys.platform == 'cygwin') \
          and str(filename)[1] == ':': # full directory for Windows
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
            config_file = 'SIREN.ini'
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
        self.adjust_re = False
        self.change_res = True
        self.details = True
        self.corrected_lcoe = True
        self.carbon_price = 0.
        self.optimise_choice = 'LCOE'
        choices = ['LCOE', 'Multi', 'Both']
        self.optimise_generations = 20
        self.optimise_mutation = 0.005
        self.optimise_population = 50
        self.optimise_stop = 0
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
                 elif key == 'generators_details':
                     if value.lower() in ['false', 'off', 'no']:
                         self.details = False
                 elif key == 'log_status':
                     if value.lower() in ['false', 'no', 'off']:
                         self.log_status = False
                 elif key == 'more_details':
                     if value.lower() in ['true', 'yes', 'on']:
                         self.more_details = True
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
        except:
            pass
        self.opt_progressbar = None
        self.floatstatus = None # status window
        self.adjustby = None
   #     self.tabs = QtGui.QTabWidget()    # Create tabs
   #     tab1 = QtGui.QWidget()
   #     tab2 = QtGui.QWidget()
   #     tab3 = QtGui.QWidget()
   #     tab4 = QtGui.QWidget()
   #     tab5 = QtGui.QWidget()
        self.grid = QtGui.QGridLayout()
        self.files = [None] * 5
        self.sheets = self.file_labels[:]
        del self.sheets[-2:]
        self.updated = False
        edit = [None, None, None]
        r = 0
        for i in range(5):
            self.grid.addWidget(QtGui.QLabel(self.file_labels[i] + ' File:'), r, 0)
            self.files[i] = ClickableQLabel()
            self.files[i].setStyleSheet("background-color: white; border: 1px inset grey; min-height: 22px; border-radius: 4px;")
            self.files[i].setText(self.ifiles[i])
            self.connect(self.files[i], QtCore.SIGNAL('clicked()'), self.fileChanged)
            self.grid.addWidget(self.files[i], r, 1, 1, 3)
            if i < D:
                r += 1
                self.grid.addWidget(QtGui.QLabel(self.file_labels[i] + ' Sheet:'), r, 0)
                self.sheets[i] = QtGui.QComboBox()
                self.sheets[i].addItem(self.isheets[i])
                self.grid.addWidget(self.sheets[i], r, 1, 1, 2)
                edit[i] = QtGui.QPushButton(self.file_labels[i], self)
                self.grid.addWidget(edit[i], r, 3)
                edit[i].clicked.connect(self.editClicked)
            r += 1
        wdth = edit[1].fontMetrics().boundingRect(edit[1].text()).width() + 9
        self.grid.addWidget(QtGui.QLabel('Carbon Price:'), r, 0)
        self.carbon = QtGui.QDoubleSpinBox()
        self.carbon.setRange(0, 200)
        self.carbon.setDecimals(2)
        try:
            self.carbon.setValue(self.carbon_price)
        except:
            self.carbon.setValue(0.)
        self.grid.addWidget(self.carbon, r, 1)
        self.carbon.valueChanged.connect(self.cpchanged)
        self.grid.addWidget(QtGui.QLabel('($/tCO2e. Use only if LCOE excludes carbon price)'), r, 2, 1, 2)
        r += 1
        self.grid.addWidget(QtGui.QLabel('Adjust Generators:'), r, 0)
        self.adjust = QtGui.QCheckBox('(check to adjust/multiply generators capacity data)', self)
        if self.adjust_re:
            self.adjust.setCheckState(QtCore.Qt.Checked)
        self.grid.addWidget(self.adjust, r, 1, 1, 3)
        r += 1
        self.grid.addWidget(QtGui.QLabel('Dispatch Order:\n(move to right\nto exclude)'), r, 0)
        self.order = ThumbListWidget(self) #QtGui.QListWidget()
      #  self.order.setDragDropMode(QtGui.QAbstractItemView.InternalMove)
        self.grid.addWidget(self.order, r, 1, 1, 2)
        self.ignore = ThumbListWidget(self) # QtGui.QListWidget()
      #  self.ignore.setDragDropMode(QtGui.QAbstractItemView.InternalMove)
        self.grid.addWidget(self.ignore, r, 3, 1, 2)
        r += 1
        self.log = QtGui.QLabel('')
        msg_palette = QtGui.QPalette()
        msg_palette.setColor(QtGui.QPalette.Foreground, QtCore.Qt.red)
        self.log.setPalette(msg_palette)
        self.grid.addWidget(self.log, r, 1, 1, 4)
        r += 1
        self.progressbar = QtGui.QProgressBar()
        self.progressbar.setMinimum(0)
        self.progressbar.setMaximum(10)
        self.progressbar.setValue(0)
        self.progressbar.setStyleSheet('QProgressBar {border: 1px solid grey; border-radius: 2px; text-align: center;}' \
                                       + 'QProgressBar::chunk { background-color: #6891c6;}')
        self.grid.addWidget(self.progressbar, r, 1, 1, 4)
        self.progressbar.setHidden(True)
        r += 1
        r += 1
        quit = QtGui.QPushButton('Done', self)
        self.grid.addWidget(quit, r, 0)
        quit.clicked.connect(self.quitClicked)
        QtGui.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
        pms = QtGui.QPushButton('Summary', self)
        self.grid.addWidget(pms, r, 1)
        pms.clicked.connect(self.pmClicked)
        pm = QtGui.QPushButton('Powermatch', self)
     #   pm.setMaximumWidth(wdth)
        self.grid.addWidget(pm, r, 2)
        pm.clicked.connect(self.pmClicked)
        opt = QtGui.QPushButton('Optimise', self)
        self.grid.addWidget(opt, r, 3)
        opt.clicked.connect(self.optClicked)
        help = QtGui.QPushButton('Help', self)
        help.setMaximumWidth(wdth)
        quit.setMaximumWidth(wdth)
        self.grid.addWidget(help, r, 5)
        help.clicked.connect(self.helpClicked)
        QtGui.QShortcut(QtGui.QKeySequence('F1'), self, self.helpClicked)
        try:
            ts = xlrd.open_workbook(self.get_filename(str(self.files[G].text())))
            ws = ts.sheet_by_name('Generators')
            self.getGenerators(ws)
            self.setOrder()
            ts.release_resources()
            del ts
        except:
            pass
        frame = QtGui.QFrame()
        frame.setLayout(self.grid)
        self.scroll = QtGui.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtGui.QVBoxLayout(self)
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
        screen = QtGui.QApplication.desktop().screenNumber(QtGui.QApplication.desktop().cursor().pos())
        centerPoint = QtGui.QApplication.desktop().availableGeometry(screen).center()
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
            newfile = str(QtGui.QFileDialog.getSaveFileName(None, 'Save ' + self.file_labels[i] + ' file',
                      curfile, 'Excel Files (*.xlsx)'))
        else:
            newfile = str(QtGui.QFileDialog.getOpenFileName(self, 'Open ' + self.file_labels[i] + ' file',
                      curfile))
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
                    self.sheets[i].addItem(str(sht))
                    if str(sht) == self.file_labels[i]:
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
                newfile = str(self.files[D].text())
                newfile = newfile.replace('data', 'results')
                newfile = newfile.replace('Data', 'Results')
                if newfile != str(self.files[D].text()):
                    self.files[R].setText(newfile)
            self.updated = True

    def helpClicked(self):
        dialog = displayobject.AnObject(QtGui.QDialog(), self.help,
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
            lines.append('carbon_price=' + str(self.carbon_price))
            for i in range(len(self.file_labels)):
                lines.append(str(self.file_labels[i].lower()) + '_file=' + str(self.files[i].text()))
            for i in range(D):
                lines.append(str(self.file_labels[i].lower()) + '_sheet=' + str(self.sheets[i].currentText()))
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
                    self.constraints[key] = Constraint(key, '<category>', 0., 1.,
                                                       1., 1., 1., 0., 1., 0., 0.)
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
                ts = xlrd.open_workbook(self.get_filename(str(self.files[it].text())))
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
            sp_pts = [2] * 11
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
                        caps = self.optimisation[key].capacities.split(' ')
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
                                              1., 1., 1., 0., 1., 0., 0.)
            return
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
            self.constraints[str(ws.cell_value(row, 0))] = Constraint(str(ws.cell_value(row, 0)),
                                     str(ws.cell_value(row, cat_col)),
                                     ws.cell_value(row, cap_col[0]), ws.cell_value(row, cap_col[1]),
                                     ws.cell_value(row, ramp_col[0]), ws.cell_value(row, ramp_col[1]),
                                     ws.cell_value(row, rec_col[0]), ws.cell_value(row, rec_col[1]),
                                     ws.cell_value(row, dis_col[0]), ws.cell_value(row, dis_col[1]),
                                     ws.cell_value(row, par_col))
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
            elif ws.cell_value(0, col) == 'Emissions':
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
            tech = str(ws.cell_value(row, 0))
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
        if self.sender().text() == 'Summary': # summary only?
            summ_only = True
        else:
            summ_only = False
        self.progressbar.setMinimum(0)
        self.progressbar.setMaximum(10)
        self.progressbar.setHidden(False)
        details = self.details
        err_msg = ''
        try:
            ts = xlrd.open_workbook(self.get_filename(str(self.files[C].text())))
            ws = ts.sheet_by_name(self.sheets[C].currentText())
            self.getConstraints(ws)
            ts.release_resources()
            del ts
        except:
            err_msg = 'Error accessing Constraints'
            self.getConstraints(None)
        try:
            ts = xlrd.open_workbook(self.get_filename(str(self.files[G].text())))
            ws = ts.sheet_by_name(self.sheets[G].currentText())
            self.getGenerators(ws)
            ts.release_resources()
            del ts
        except:
            if err_msg != '':
                err_msg += ' and Generators'
            else:
                err_msg = 'Error accessing Generators.'
            self.getGenerators(None)
        if err_msg != '':
            self.setStatus(err_msg)
        start_time = time.time()
        re_capacities = [0.] * len(tech_names)
        pm_data_file = self.get_filename(str(self.files[D].text()))
        re_capacity = 0.
        data = []
        load = []
        shortfall = []
        if pm_data_file[-4:] == '.xls': #xls format
            xlsx = False
            details = False
            ts = xlrd.open_workbook(pm_data_file)
            ws = ts.sheet_by_index(0)
            if ws.cell_value(0, 0) != 'Hourly Shortfall Table' \
              or ws.cell_value(0, 4) != 'Generation Summary Table':
                self.setStatus('not a Powerbalance spreadsheet')
                self.progressbar.setHidden(True)
                return
            for row in range(20):
                if ws.cell_value(row, 5) == 'Total':
                    re_capacity = ws.cell_value(row, 6)
                    re_generation = ws.cell_value(row, 8)
                    break
            for row in range(2, ws.nrows):
                shortfall.append(ws.cell_value(row, 2))
        else: # xlsx format
            xlsx = True
            shortfall = [0.] * 8760
            if details:
                cols = []
            ts = oxl.load_workbook(pm_data_file)
            ws = ts.active
            top_row = ws.max_row - 8760
            if ws.cell(row=top_row, column=1).value != 'Hour' or ws.cell(row=top_row, column=2).value != 'Period':
                self.setStatus('not a Powermatch data spreadsheet')
                self.progressbar.setHidden(True)
                return
            typ_row = top_row - 1
            while typ_row > 0:
                if ws.cell(row=typ_row, column=3).value in tech_names:
                    break
                typ_row -= 1
            else:
                self.setStatus('no suitable data')
                return
            icap_row = typ_row + 1
            while icap_row < top_row:
                if ws.cell(row=icap_row, column=1).value == 'Capacity (MW)':
                    break
                icap_row += 1
            else:
                self.setStatus('no capacity data')
                return
       #     adjustby = None
            if self.adjust.isChecked():
                adjustin = []
                for col in range(3, ws.max_column + 1):
                    try:
                        valu = ws.cell(row=typ_row, column=col).value.replace('-','')
                        i = tech_names.index(valu)
                    except:
                        break
                    if valu == 'Load':
                        adjustin.append([tech_names[i], 0])
                    else:
                        try:
                            adjustin.append([tech_names[i], float(ws.cell(row=icap_row, column=col).value)])
                        except:
                            pass
                for i in range(self.order.count()):
                    if self.generators[self.order.item(i).text()].capacity > 0:
                        adjustin.append([self.order.item(i).text(),
                                        self.generators[self.order.item(i).text()].capacity])
                adjust = Adjustments(adjustin, self.adjustby)
                adjust.exec_()
                self.adjustby = adjust.getValues()
                if self.adjustby is None:
                    self.setStatus('Execution aborted.')
                    self.progressbar.setHidden(True)
                    return
            load_col = -1
            det_col = 3
            self.progressbar.setValue(1)
            for col in range(3, ws.max_column + 1):
                try:
                    valu = ws.cell(row=typ_row, column=col).value.replace('-','')
                    i = tech_names.index(valu)
                except:
                    break # ?? or continue??
                if tech_names[i] != 'Load':
                    try:
                        if ws.cell(row=icap_row, column=col).value <= 0:
                            continue
                    except:
                        continue
                data.append([])
                try:
                    multiplier = self.adjustby[tech_names[i]]
                except:
                    multiplier = 1.
                if details:
                    cols.append(i)
                    try:
                        re_capacities[i] = ws.cell(row=icap_row, column=col).value * multiplier
                    except:
                        pass
                try:
                    re_capacity += ws.cell(row=icap_row, column=col).value * multiplier
                except:
                    pass
                for row in range(top_row + 1, ws.max_row + 1):
                    data[-1].append(ws.cell(row=row, column=col).value * multiplier)
                    try:
                        if tech_names[i] == 'Load':
                            load_col = len(data) - 1
                            shortfall[row - top_row - 1] += ws.cell(row=row, column=col).value * multiplier
                        else:
                            shortfall[row - top_row - 1] -= ws.cell(row=row, column=col).value * multiplier
                    except:
                        pass
            ts.close()
        if str(self.files[R].text()) == '':
            i = pm_data_file.find('/')
            if i >= 0:
                data_file = pm_data_file[i + 1:]
            else:
                data_file = pm_data_file
            data_file = data_file.replace('data', 'results')
        else:
            data_file = self.get_filename(str(self.files[R].text()))
        if not xlsx:
            data_file += 'x'
        self.progressbar.setValue(2)
        headers = ['Facility', 'Capacity (MW)', 'Subtotal (MWh)', 'CF', 'Cost ($)', 'LCOE ($/MWh)', 'Emissions (tCO2e)']
        sp_cols = []
        sp_cap = []
        if summ_only: # summary only?
            sp_data = []
            sp_gen = []
            sp_load = 0.
            for i in cols:
                if tech_names[i] == 'Load':
                    sp_load = sum(data[load_col])
                    continue
                sp_data.append([tech_names[i], re_capacities[i], 0., '', '', '', ''])
                for g in data[len(sp_data)]:
                    sp_data[-1][2] += g
                # if ignore not used
        else: # normal
            ds = oxl.Workbook()
            ns = ds.active
            ns.title = 'Detail'
            ss = ds.create_sheet('Summary', 0)
            re_sum = '=('
            cap_row = 1
            ns.cell(row=cap_row, column=2).value = headers[1]
            ss.cell(row=3, column=1).value = headers[0]
            ss.cell(row=3, column=2).value = headers[1]
            ini_row = 2
            ns.cell(row=ini_row, column=2).value = 'Initial Capacity'
            sum_row = 3
            ns.cell(row=sum_row, column=2).value = headers[2]
            ss.cell(row=3, column=3).value = headers[2]
            cf_row = 4
            ns.cell(row=cf_row, column=2).value = headers[3]
            ss.cell(row=3, column=4).value = headers[3]
            cost_row = 5
            ns.cell(row=cost_row, column=2).value = headers[4]
            ss.cell(row=3, column=5).value = headers[4]
            lcoe_row = 6
            ns.cell(row=lcoe_row, column=2).value = headers[5]
            ss.cell(row=3, column=6).value = headers[5]
            emi_row = 7
            ns.cell(row=emi_row, column=2).value = headers[6]
            ss.cell(row=3, column=7).value = headers[6]
            ss_row = 3
            fall_row = 8
            ns.cell(row=fall_row, column=2).value = 'Shortfall periods'
            what_row = 9
            hrows = 10
            ns.cell(row=what_row, column=1).value = 'Hour'
            ns.cell(row=what_row, column=2).value = 'Period'
            ns.cell(row=what_row, column=3).value = 'Load'
            ns.cell(row=sum_row, column=3).value = '=SUM(' + ss_col(3) + str(hrows) + \
                                                   ':' + ss_col(3) + str(hrows + 8759) + ')'
            o = 4
            col = 3
            if details:
                for i in cols:
                    if tech_names[i] == 'Load':
                        continue
                    col += 1
                    sp_cols.append(tech_names[i])
                    sp_cap.append(re_capacities[i])
                    ss_row += 1
                    ns.cell(row=what_row, column=col).value = tech_names[i]
                    ss.cell(row=ss_row, column=1).value = '=Detail!' + ss_col(col) + str(what_row)
                    ns.cell(row=cap_row, column=col).value = re_capacities[i]
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
                    if tech_names[i] not in self.generators:
                        continue
                    if self.generators[tech_names[i]].lcoe > 0:
                        ns.cell(row=cost_row, column=col).value = self.generators[tech_names[i]].lcoe * \
                                self.generators[tech_names[i]].lcoe_cf * 8760 * re_capacities[i]
                        ns.cell(row=cost_row, column=col).number_format = '$#,##0'
                        ss.cell(row=ss_row, column=5).value = '=Detail!' + ss_col(col) + str(cost_row)
                        ss.cell(row=ss_row, column=5).number_format = '$#,##0'
                        ns.cell(row=lcoe_row, column=col).value = '=IF(AND(' + ss_col(col) + str(cf_row) + '>0,' \
                                + ss_col(col) + str(cap_row) + '>0),' + ss_col(col) + str(cost_row) + '/8760/' \
                                + ss_col(col) + str(cf_row) +'/' + ss_col(col) + str(cap_row) + ',"")'
                        ns.cell(row=lcoe_row, column=col).number_format = '$#,##0.00'
                        ss.cell(row=ss_row, column=6).value = '=Detail!' + ss_col(col) + str(lcoe_row)
                        ss.cell(row=ss_row, column=6).number_format = '$#,##0.00'
                    if self.generators[tech_names[i]].emissions > 0:
                        ns.cell(row=emi_row, column=col).value = '=' + ss_col(col) + str(sum_row) \
                                + '*' + str(self.generators[tech_names[i]].emissions)
                        ns.cell(row=emi_row, column=col).number_format = '#,##0'
                        ss.cell(row=ss_row, column=7).value = '=Detail!' + ss_col(col) + str(emi_row)
                        ss.cell(row=ss_row, column=7).number_format = '#,##0'
                shrt_col = col + 1
            else:
                shrt_col = 5
                sp_cols.append('Renewable')
                sp_cap.append(re_capacity)
                ns.cell(row=what_row, column=4).value = 'Renewable'
                ns.cell(row=cap_row, column=4).value = re_capacity
                ns.cell(row=cap_row, column=4).number_format = '#,##0.00'
                if xlsx:
                    ns.cell(row=sum_row, column=4).value = '=SUM(D' + str(hrows) + ':D' + str(hrows + 8759) + ')'
                else:
                    ns.cell(row=sum_row, column=4).value = re_generation
                ns.cell(row=sum_row, column=4).number_format = '#,##0'
                ns.cell(row=cf_row, column=4).value = '=IF(D1>0,D3/D1/8760."")'
                ns.cell(row=cf_row, column=4).number_format = '#,##0.00'
                ss_row += 1
                ss.cell(row=ss_row, column=1).value = '=Detail!D' + str(what_row)
                ss.cell(row=ss_row, column=2).value = '=Detail!D' + str(cap_row)
                ss.cell(row=ss_row, column=2).number_format = '#,##0.00'
                ss.cell(row=ss_row, column=3).value = '=Detail!D' + str(sum_row)
                ss.cell(row=ss_row, column=3).number_format = '#,##0'
                ss.cell(row=ss_row, column=4).value = '=Detail!D' + str(cf_row)
                ss.cell(row=ss_row, column=4).number_format = '#,##0.00'
            ns.cell(row=fall_row, column=shrt_col).value = '=COUNTIF(' + ss_col(shrt_col) \
                            + str(hrows) + ':' + ss_col(shrt_col) + str(hrows + 8759) + ',">0")'
            ns.cell(row=fall_row, column=shrt_col).number_format = '#,##0'
            ns.cell(row=what_row, column=shrt_col).value = 'Shortfall'
            for col in range(3, shrt_col + 1):
                ns.cell(row=what_row, column=col).alignment = oxl.styles.Alignment(wrap_text=True,
                        vertical='bottom', horizontal='center')
            for row in range(hrows, 8760 + hrows):
                if xlsx:
                    ns.cell(row=row, column=1).value = ws.cell(row=top_row + row - hrows + 1, column=1).value
                    ns.cell(row=row, column=2).value = ws.cell(row=top_row + row - hrows + 1, column=2).value
                    ns.cell(row=row, column=3).value = round(data[load_col][row - hrows], 2)
                    if not details:
                        ns.cell(row=row, column=4).value = round(data[load_col][row - hrows] - shortfall[row - hrows], 2)
                else:
                    ns.cell(row=row, column=1).value = ws.cell_value(row - hrows + 2, 0)
                    ns.cell(row=row, column=2).value = ws.cell_value(row - hrows + 2, 1)
                ns.cell(row=row, column=shrt_col).value = round(shortfall[row - hrows], 2)
                for col in range(3, shrt_col + 1):
                    ns.cell(row=row, column=col).number_format = '#,##0.00'
            if details:
                col = 3
                for i in range(len(data)):
                    if cols[i] == load_col:
                        continue
                    col += 1
                    for row in range(hrows, 8760 + hrows):
                        ns.cell(row=row, column=col).value = data[i][row - hrows]
            col = shrt_col + 1
        order = [] #'Storage', 'Biomass', 'PHS', 'Gas', 'CCG1', 'Other', 'Coal']
        for itm in range(self.order.count()):
            order.append(str(self.order.item(itm).text()))
        #storage? = []
        self.progressbar.setValue(3)
        for gen in order:
            try:
                capacity = self.generators[gen].capacity * self.adjustby[gen]
            except:
                capacity = self.generators[gen].capacity
            if self.generators[gen].constraint in self.constraints and \
              self.constraints[self.generators[gen].constraint].category == 'Storage': # storage
                storage = [0., 0., 0., 0.] # capacity, initial, min level, max drain
                storage[0] = capacity
                if not summ_only:
                    ns.cell(row=cap_row, column=col).value = round(capacity, 2)
                    ns.cell(row=cap_row, column=col).number_format = '#,##0.00'
                storage[1] = self.generators[gen].initial
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
                storage_carry = storage[1] # self.generators[gen].initial
                if not summ_only:
                    ns.cell(row=ini_row, column=col).value = round(storage_carry, 2)
                    ns.cell(row=ini_row, column=col).number_format = '#,##0.00'
                storage_bal = []
                storage_can = 0.
                for row in range(8760):
                    if storage_carry > 0:
                        storage_carry = storage_carry * (1 - parasite)
                    storage_loss = 0.
                    if shortfall[row] < 0:  # excess generation
                        can_use = - (storage[0] - storage_carry) * (1 / (1 - recharge[1]))
                        if can_use < 0: # can use some
                            if shortfall[row] > can_use:
                                can_use = shortfall[row]
                            if can_use < - recharge[0] * (1 / (1 - recharge[1])):
                                can_use = - recharge[0]
                        else:
                            can_use = 0.
                        storage_carry -= (can_use * (1 - recharge[1]))
                        shortfall[row] -= can_use
                    else: # shortfall
                        can_use = shortfall[row] * (1 / (1 - discharge[1]))
                        if can_use > storage_carry - storage[2]:
                            can_use = storage_carry - storage[2]
                        if can_use > 0:
                            storage_loss = can_use * discharge[1]
                            storage_carry -= can_use
                            shortfall[row] -= (can_use - storage_loss)
                            if storage_carry < 0:
                                storage_carry = 0
                        else:
                            can_use = 0.
                    storage_bal.append(storage_carry)
                    if not summ_only:
                        ns.cell(row=row + hrows, column=col).value = round(can_use, 2)
                        ns.cell(row=row + hrows, column=col + 1).value = round(storage_carry, 2)
                        ns.cell(row=row + hrows, column=col + 2).value = round(shortfall[row], 2)
                        ns.cell(row=row + hrows, column=col).number_format = '#,##0.00'
                        ns.cell(row=row + hrows, column=col + 1).number_format = '#,##0.00'
                        ns.cell(row=row + hrows, column=col + 2).number_format = '#,##0.00'
                    else:
                        if can_use > 0:
                            storage_can += can_use
                if not summ_only:
                    ns.cell(row=sum_row, column=col).value = '=SUMIF(' + ss_col(col) + \
                            str(hrows) + ':' + ss_col(col) + str(hrows + 8759) + ',">0")'
                    ns.cell(row=sum_row, column=col).number_format = '#,##0'
                    ns.cell(row=cf_row, column=col).value = '=IF(' + ss_col(col) + '1>0,' + \
                                                    ss_col(col) + '3/' + ss_col(col) + '1/8760,"")'
                    ns.cell(row=cf_row, column=col).number_format = '#,##0.00'
                    col += 3
                else:
                    sp_data.append([gen, storage[0], storage_can, '', '', '', ''])
            else:
                if not summ_only:
                    ns.cell(row=cap_row, column=col).value = round(capacity, 2)
                    ns.cell(row=cap_row, column=col).number_format = '#,##0.00'
                    for row in range(8760):
                        if shortfall[row] >= 0: # shortfall?
                            if shortfall[row] >= capacity:
                                shortfall[row] = shortfall[row] - capacity
                                ns.cell(row=row + hrows, column=col).value = round(capacity, 2)
                            else:
                                ns.cell(row=row + hrows, column=col).value = round(shortfall[row], 2)
                                shortfall[row] = 0
                        else:
                            ns.cell(row=row + hrows, column=col).value = 0
                        ns.cell(row=row + hrows, column=col + 1).value = round(shortfall[row], 2)
                        ns.cell(row=row + hrows, column=col).number_format = '#,##0.00'
                        ns.cell(row=row + hrows, column=col + 1).number_format = '#,##0.00'
                    ns.cell(row=sum_row, column=col).value = '=SUM(' + ss_col(col) + str(hrows) + \
                            ':' + ss_col(col) + str(hrows + 8759) + ')'
                    ns.cell(row=sum_row, column=col).number_format = '#,##0'
                    ns.cell(row=cf_row, column=col).value = '=IF(' + ss_col(col) + '1>0,' + \
                                                ss_col(col) + '3/' + ss_col(col) + '1/8760,"")'
                    ns.cell(row=cf_row, column=col).number_format = '#,##0.00'
                    col += 2
                else:
                    gen_can = 0.
                    for row in range(8760):
                        if shortfall[row] >= 0: # shortfall?
                            if shortfall[row] >= capacity:
                                shortfall[row] = shortfall[row] - capacity
                                gen_can += capacity
                            else:
                                gen_can += shortfall[row]
                                shortfall[row] = 0
                    sp_data.append([gen, capacity, gen_can, '', '', '', ''])
        self.progressbar.setValue(4)
        if summ_only:
            cap_sum = 0.
            gen_sum = 0.
            re_sum = 0.
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
                if self.generators[gen].lcoe > 0:
                    sp_data[sp][4] = self.generators[gen].lcoe * self.generators[gen].lcoe_cf * 8760 * sp_data[sp][1]
                    if sp_data[sp][1] > 0 and sp_data[sp][3] > 0:
                        sp_data[sp][5] = sp_data[sp][4] / 8760 / sp_data[sp][3] / sp_data[sp][1]
                    cost_sum += sp_data[sp][4]
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
            sp_data.append(['Total', cap_sum, gen_sum, cs, cost_sum, gs, co2_sum])
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
            sf_sums = [0., 0., 0.]
            for sf in range(len(shortfall)):
                if shortfall[sf] > 0:
                    sf_sums[0] += shortfall[sf]
                    sf_sums[2] += data[0][sf]
                else:
                    sf_sums[1] += shortfall[sf]
                    sf_sums[2] += data[0][sf]
            sp_data.append(' ')
            sp_data.append('Load Analysis')
            pct = '{:.1%})'.format((sf_sums[2] - sf_sums[0]) / sp_load)
            sp_data.append(['Load met (' + pct, '', sf_sums[2] - sf_sums[0]])
            pct = '{:.1%})'.format(sf_sums[0] / sp_load)
            sp_data.append(['Shortfall (' + pct, '', sf_sums[0]])
            sp_data.append(['Total Load', '', sp_load])
            pct = '{:.1%})'.format( -sf_sums[1] / sp_load)
            sp_data.append(['Surplus (' + pct, '', -sf_sums[1]])
            adjusted = False
            if self.adjustby is not None:
                for key, value in iter(sorted(self.adjustby.items())):
                    if value != 1:
                        if not adjusted:
                            adjusted = True
                            sp_data.append('Generators Adjustments:')
                        sp_data.append([key, value])
            list(map(list, list(zip(*sp_data))))
            sp_pts = [0, 2, 0, 2, 0, 2, 0]
            self.setStatus(self.sender().text() + ' completed')
            dialog = displaytable.Table(sp_data, title=self.sender().text(), fields=headers,
                     save_folder=self.scenarios, sortby='', decpts=sp_pts)
            dialog.exec_()
            self.progressbar.setValue(10)
            self.progressbar.setHidden(True)
            self.progressbar.setValue(0)
            return
        for column_cells in ns.columns:
            length = 0
            value = ''
            row = 0
            for cell in column_cells:
                if cell.row >= 0:
                    if len(str(cell.value)) > length:
                        length = len(str(cell.value))
                        value = cell.value
                        row = cell.row
            try:
                ns.column_dimensions[column_cells[0].column].width = max(length, 10)
            except:
                ns.column_dimensions[ss_col(column_cells[0].column)].width = max(length, 10)
        ns.column_dimensions['A'].width = 6
        col = shrt_col + 1
        for gen in order:
            ss_row += 1
            ns.cell(row=what_row, column=col).value = gen
            ss.cell(row=ss_row, column=1).value = '=Detail!' + ss_col(col) + str(what_row)
            ss.cell(row=ss_row, column=2).value = '=Detail!' + ss_col(col) + str(cap_row)
            ss.cell(row=ss_row, column=2).number_format = '#,##0.00'
            ss.cell(row=ss_row, column=3).value = '=Detail!' + ss_col(col) + str(sum_row)
            ss.cell(row=ss_row, column=3).number_format = '#,##0'
            ss.cell(row=ss_row, column=4).value = '=Detail!' + ss_col(col) + str(cf_row)
            ss.cell(row=ss_row, column=4).number_format = '#,##0.00'
            if self.generators[gen].lcoe > 0:
                try:
                    capacity = self.generators[gen].capacity * self.adjustby[gen]
                except:
                   capacity = self.generators[gen].capacity
                ns.cell(row=cost_row, column=col).value = self.generators[gen].lcoe * \
                        self.generators[gen].lcoe_cf * 8760 * capacity
                ns.cell(row=cost_row, column=col).number_format = '$#,##0'
                ss.cell(row=ss_row, column=5).value = '=Detail!' + ss_col(col) + str(cost_row)
                ss.cell(row=ss_row, column=5).number_format = '$#,##0'
                ns.cell(row=lcoe_row, column=col).value = '=IF(AND(' + ss_col(col) + str(cf_row) + '>0,' \
                            + ss_col(col) + str(cap_row) + '>0),' + ss_col(col) + str(cost_row) + '/8760/' \
                            + ss_col(col) + str(cf_row) + '/' + ss_col(col) + str(cap_row)+  ',"")'
                ns.cell(row=lcoe_row, column=col).number_format = '$#,##0.00'
                ss.cell(row=ss_row, column=6).value = '=Detail!' + ss_col(col) + str(lcoe_row)
                ss.cell(row=ss_row, column=6).number_format = '$#,##0.00'
            if self.generators[gen].emissions > 0:
                ns.cell(row=emi_row, column=col).value = '=' + ss_col(col) + str(sum_row) \
                        + '*' + str(self.generators[gen].emissions)
                ns.cell(row=emi_row, column=col).number_format = '#,##0'
                ss.cell(row=ss_row, column=7).value = '=Detail!' + ss_col(col) + str(emi_row)
                ss.cell(row=ss_row, column=7).number_format = '#,##0'
            ns.cell(row=what_row, column=col).alignment = oxl.styles.Alignment(wrap_text=True,
                    vertical='bottom', horizontal='center')
            ns.cell(row=what_row, column=col + 1).alignment = oxl.styles.Alignment(wrap_text=True,
                    vertical='bottom', horizontal='center')
            if self.constraints[self.generators[gen].constraint].category == 'Storage':
                ns.cell(row=what_row, column=col + 1).value = gen + '\nBalance'
                ns.cell(row=what_row, column=col + 2).value = 'After\n' + gen
                ns.cell(row=what_row, column=col + 2).alignment = oxl.styles.Alignment(wrap_text=True,
                        vertical='bottom', horizontal='center')
                ns.cell(row=fall_row, column=col + 2).value = '=COUNTIF(' + ss_col(col + 2) \
                        + str(hrows) + ':' + ss_col(col + 2) + str(hrows + 8759) + ',">0")'
                ns.cell(row=fall_row, column=col + 2).number_format = '#,##0'
                col += 3
            else:
                ns.cell(row=what_row, column=col + 1).value = 'After\n' + gen
                ns.cell(row=fall_row, column=col + 1).value = '=COUNTIF(' + ss_col(col + 1) \
                        + str(hrows) + ':' + ss_col(col + 1) + str(hrows + 8759) + ',">0")'
                ns.cell(row=fall_row, column=col + 1).number_format = '#,##0'
                col += 2
        self.progressbar.setValue(6)
        ns.row_dimensions[what_row].height = 30
        ns.freeze_panes = 'C' + str(hrows)
        ns.activeCell = 'C' + str(hrows)
        ss.cell(row=1, column=1).value = 'Powermatch - Summary'
        bold = oxl.styles.Font(bold=True, name=ss.cell(row=1, column=1).font.name)
        ss.cell(row=1, column=1).font = bold
        ss_row +=1
        for col in range(1, 8):
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
            ss.cell(row=ss_row, column=1).value = 'Corrected LCOE)'
            ss.cell(row=ss_row, column=1).font = bold
        lcoe_row = ss_row
        for column_cells in ss.columns:
            length = 0
            value = ''
            row = 0
            for cell in column_cells:
                if len(str(cell.value)) > length:
                    length = len(str(cell.value))
                    value = cell.value
                    row = cell.row
            try:
                ss.column_dimensions[column_cells[0].column].width = length * 1.15
            except:
                ss.column_dimensions[ss_col(column_cells[0].column)].width = length * 1.15
        ss.column_dimensions['D'].width = 7
        ss.column_dimensions['E'].width = 18
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
            ss.cell(row=ss_row, column=1).value = 'Carbon Cost ($' + cp + ')'
            ss.cell(row=ss_row, column=5).value = '=G' + str(ss_row - r) + '*' + \
                    str(self.carbon_price)
            ss.cell(row=ss_row, column= 5).number_format = '$#,##0'
            if not self.corrected_lcoe:
                ss.cell(row=ss_row, column=6).value = '=(E' + str(ss_row - r) + \
                        '+E'  + str(ss_row) + ')/C' + str(ss_row - r)
                ss.cell(row=ss_row, column=6).number_format = '$#,##0.00'
            r += 1
        ss_row += 1
        ss.cell(row=ss_row, column=1).value = 'RE %age'
        ss.cell(row=ss_row, column=2).value = re_sum[:-1] + ')/C' + str(ss_row - r)
        ss.cell(row=ss_row, column=2).number_format = '#,##0.0%'
        ss_row += 2
        ss.cell(row=ss_row, column=1).value = 'Load Analysis'
        ss.cell(row=ss_row, column=1).font = bold
        ss_row += 1
        ss.cell(row=ss_row, column=1).value = 'Load met'
        ss.cell(row=ss_row, column=3).value = '=SUMIF(Detail!' + last_col + str(hrows) + ':Detail!' \
            + last_col + str(hrows + 8759) + ',"<=0",Detail!C' + str(hrows) + ':Detail!C' \
            + str(hrows + 8759) + ')+SUMIF(Detail!' + last_col + str(hrows) + ':Detail!' \
            + last_col + str(hrows + 8759) + ',">0",Detail!C' + str(hrows) + ':Detail!C' \
            + str(hrows + 8759) + ')-C' + str(ss_row + 1)
        ss.cell(row=ss_row, column=3).number_format = '#,##0'
        ss.cell(row=ss_row, column=4).value = '=C' + str(ss_row) + '/C' + str(ss_row + 2)
        ss.cell(row=ss_row, column=4).number_format = '#,##0.0%'
        ss_row += 1
        ss.cell(row=ss_row, column=1).value = 'Shortfall'
        ss.cell(row=ss_row, column=3).value = '=SUMIF(Detail!' + last_col + str(hrows) + ':Detail!' \
            + last_col + str(hrows + 8759) + ',">0",Detail!' + last_col + str(hrows) + ':Detail!' \
            + last_col + str(hrows + 8759) + ')'
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
            ss.cell(row=lcoe_row + 1, column=6).value = '=(E' + str(lcoe_row - 1) + \
                    '+E'  + str(lcoe_row + 1) + ')/C' + str(ss_row)
            ss.cell(row=lcoe_row + 1, column=6).number_format = '$#,##0.00'
        ss_row += 1
        ss.cell(row=ss_row, column=1).value = 'Surplus'
        ss.cell(row=ss_row, column=3).value = '=-SUMIF(Detail!' + last_col + str(hrows) + ':Detail!' \
            + last_col + str(hrows + 8759) + ',"<0",Detail!' + last_col + str(hrows) \
            + ':Detail!' + last_col + str(hrows + 8759) + ')'
        ss.cell(row=ss_row, column=3).number_format = '#,##0'
        ss.cell(row=ss_row, column=4).value = '=C' + str(ss_row) + '/C' + str(ss_row - 1)
        ss.cell(row=ss_row, column=4).number_format = '#,##0.0%'
        ss_row += 2
        ss.cell(row=ss_row, column=1).value = 'Data sources:'
        ss.cell(row=ss_row, column=1).font = bold
        ss_row += 1
        ss.cell(row=ss_row, column=1).value = 'Scenarios folder'
        ss.cell(row=ss_row, column=2).value = self.scenarios
        ss.merge_cells('B' + str(ss_row) + ':G' + str(ss_row))
        ss_row += 1
        ss.cell(row=ss_row, column=1).value = 'Powermatch data file'
        if pm_data_file[: len(self.scenarios)] == self.scenarios:
            pm_data_file = pm_data_file[len(self.scenarios):]
        ss.cell(row=ss_row, column=2).value = pm_data_file
        ss.merge_cells('B' + str(ss_row) + ':G' + str(ss_row))
        ss_row += 1
        ss.cell(row=ss_row, column=1).value = 'Constraints worksheet'
        ss.cell(row=ss_row, column=2).value = str(self.files[C].text()) \
               + '.' + str(self.sheets[C].currentText())
        ss.merge_cells('B' + str(ss_row) + ':G' + str(ss_row))
        ss_row += 1
        ss.cell(row=ss_row, column=1).value = 'Facility worksheet'
        ss.cell(row=ss_row, column=2).value = str(self.files[G].text()) \
               + '.' + str(self.sheets[G].currentText())
        ss.merge_cells('B' + str(ss_row) + ':G' + str(ss_row))
        self.progressbar.setValue(7)
        try:
            if self.adjustby is not None:
                adjusted = ''
                for key, value in iter(sorted(self.adjustby.items())):
                    if value != 1:
                        adjusted += key + ': ' + str(value) + '; '
                if len(adjusted) > 0:
                    ss_row += 1
                    ss.cell(row=ss_row, column=1).value = 'Renewable inputs adjusted'
                    ss.cell(row=ss_row, column=2).value = adjusted[:-2]
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
            self.connect(self.opt_progressbar, QtCore.SIGNAL('progress'), self.opt_progressbar.progress)
            self.connect(self.opt_progressbar, QtCore.SIGNAL('range'), self.opt_progressbar.range)
            self.opt_progressbar.show()
            self.opt_progressbar.setVisible(False)
            self.activateWindow()
        else:
            self.opt_progressbar.range(0, maximum, msg=msg)

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
            self.connect(self.floatstatus, QtCore.SIGNAL('log'), self.floatstatus.log)
            self.floatstatus.show()
            self.activateWindow()

    def setStatus(self, text):
        if self.log.text == text:
            return
        self.log.setText(text)
        if text == '':
            return
        if self.floatstatus and self.log_status:
            self.floatstatus.emit(QtCore.SIGNAL('log'), text)
            QtGui.QApplication.processEvents()

    @QtCore.pyqtSlot(str)
    def getStatus(self, text):
        if text == 'goodbye':
            self.floatstatus = None

    def exit(self):
        self.updated = False
        if self.floatstatus is not None:
            self.floatstatus.exit()
        self.close()

    def optClicked(self):

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
                        fighter_fitness[f] += (arg[f] - min1) / (max1 - min1)
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
            population[random_mutation_boolean] = np.logical_not(population[random_mutation_boolean])
            # Return mutation population
            return population

        def calculate_fitness(population):
            lcoe_fitness_scores = [] # scores = LCOE values
            multi_fitness_scores = [] # scores = multi-variable weight
            multi_values = [] # values for each of the six variables
            for chromosome in population:
                op_data = []
                shortfall = op_load[:]
                # now get random amount of generation per technology (both RE and non-RE)
                for gen, value in dict_order.items():
                    capacity = dict_order[gen][3]
                    for c in range(value[1], value[2]):
                        if chromosome[c]:
                            capacity = capacity + capacities[c]
                    self.adjustby[gen] = capacity / dict_order[gen][0]
                    if gen in orig_tech.keys():
                        op_data.append([gen, capacity, 0., '', '', '', ''])
                        adjust = capacity / dict_order[gen][0]
                        for h in range(len(shortfall)):
                            g = orig_tech[gen][h] * adjust
                            op_data[-1][2] += g
                            shortfall[h] = shortfall[h] - g
                        continue
                    if self.generators[gen].constraint in self.constraints and \
                      self.constraints[self.generators[gen].constraint].category == 'Storage': # storage
                        storage = [0., 0., 0., 0.] # capacity, initial, min level, max drain
                        storage[0] = capacity
                        storage[1] = self.generators[gen].initial
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
                        storage_carry = storage[1] # self.generators[gen].initial
                        storage_bal = []
                        storage_can = 0.
                        for row in range(8760):
                            if storage_carry > 0:
                                storage_carry = storage_carry * (1 - parasite)
                            storage_loss = 0.
                            if shortfall[row] < 0:  # excess generation
                                can_use = - (storage[0] - storage_carry) * (1 / (1 - recharge[1]))
                                if can_use < 0: # can use some
                                    if shortfall[row] > can_use:
                                        can_use = shortfall[row]
                                    if can_use < - recharge[0] * (1 / (1 - recharge[1])):
                                        can_use = - recharge[0]
                                else:
                                    can_use = 0.
                                storage_carry -= (can_use * (1 - recharge[1]))
                                shortfall[row] -= can_use
                            else: # shortfall
                                can_use = shortfall[row] * (1 / (1 - discharge[1]))
                                if can_use > storage_carry - storage[2]:
                                    can_use = storage_carry - storage[2]
                                if can_use > 0:
                                    storage_loss = can_use * discharge[1]
                                    storage_carry -= can_use
                                    shortfall[row] -= (can_use - storage_loss)
                                    if storage_carry < 0:
                                        storage_carry = 0
                                else:
                                    can_use = 0.
                            storage_bal.append(storage_carry)
                            if can_use > 0:
                                storage_can += can_use
                        op_data.append([gen, storage[0], storage_can, '', '', '', ''])
                    else:
                        gen_max_cap = 0
                        gen_can = 0.
                        if capacity > 0:
                            for row in range(8760):
                                if shortfall[row] >= 0: # shortfall?
                                    if shortfall[row] >= capacity:
                                        shortfall[row] = shortfall[row] - capacity
                                        gen_can += capacity
                                        gen_max_cap = capacity
                                    else:
                                        gen_can += shortfall[row]
                                        if shortfall[row] > gen_max_cap:
                                            gen_max_cap = shortfall[row]
                                        shortfall[row] = 0
                        op_data.append([gen, capacity, gen_can, '', '', '', ''])
                cap_sum = 0.
                gen_sum = 0.
                re_sum = 0.
                cost_sum = 0.
                co2_sum = 0.
                for sp in range(len(op_data)):
                    if op_data[sp][1] > 0:
                        cap_sum += op_data[sp][1]
                        op_data[sp][3] = op_data[sp][2] / op_data[sp][1] / 8760
                    gen_sum += op_data[sp][2]
                    gen = op_data[sp][0]
                    if gen in orig_tech.keys():
                        re_sum += op_data[sp][2]
                    if gen not in self.generators:
                        continue
                    if self.generators[gen].lcoe > 0:
                        op_data[sp][4] = self.generators[gen].lcoe * self.generators[gen].lcoe_cf * 8760 * op_data[sp][1]
                        if op_data[sp][1] > 0 and op_data[sp][3] > 0:
                            op_data[sp][5] = op_data[sp][4] / 8760 / op_data[sp][3] / op_data[sp][1]
                        cost_sum += op_data[sp][4]
                    if self.generators[gen].emissions > 0:
                        op_data[sp][6] = op_data[sp][2] * self.generators[gen].emissions
                        co2_sum += op_data[sp][6]
                if cap_sum > 0:
                    cs = gen_sum / cap_sum / 8760
                else:
                    cs = ''
                if gen_sum > 0:
                    gs = cost_sum / gen_sum
                    gsw = cost_sum / op_load_tot # corrected LCOE
                else:
                    gs = ''
                    gsw = ''
                sf_sums = [0., 0., 0.]
                for sf in range(len(shortfall)):
                    if shortfall[sf] > 0:
                        sf_sums[0] += shortfall[sf]
                        sf_sums[2] += op_load[sf]
                    else:
                        sf_sums[1] += shortfall[sf]
                        sf_sums[2] += op_load[sf]
                if (sf_sums[2] - sf_sums[0]) / op_load_tot < 1:
                    lcoe_fitness_scores.append(200)
                elif self.corrected_lcoe:
                    lcoe_fitness_scores.append(gsw) # target is corrected lcoe
                else:
                    lcoe_fitness_scores.append(gs)
                multi_values.append({'lcoe': lcoe_fitness_scores[-1], #lcoe. lower better
                    'load_pct': (sf_sums[2] - sf_sums[0]) / op_load_tot, #load met. 100% better
                    'surplus_pct': -sf_sums[1] / op_load_tot, #surplus. lower better
                    're_pct': re_sum / gen_sum, # RE pct. higher better
                    'cost': cost_sum, # cost. lower better
                    'co2': co2_sum}) # CO2. lower better
                multi_fitness_scores.append(calc_weight(multi_values[-1]))
            if len(population) == 1: # return the table for best chromosome
                op_data.append(['Total', cap_sum, gen_sum, cs, cost_sum, gs, co2_sum])
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
                op_data.append(['RE %age', round(re_sum * 100. / gen_sum, 1)])
                op_data.append(' ')
                op_data.append('Load Analysis')
                pct = '{:.1%})'.format((sf_sums[2] - sf_sums[0]) / op_load_tot)
                op_data.append(['Load met (' + pct, '', sf_sums[2] - sf_sums[0],])
                pct = '{:.1%})'.format(sf_sums[0] / op_load_tot)
                op_data.append(['Shortfall (' + pct, '', sf_sums[0]])
                op_data.append(['Total Load', '', op_load_tot])
                pct = '{:.1%})'.format( -sf_sums[1] / op_load_tot)
                op_data.append(['Surplus (' + pct, '', -sf_sums[1]])
                return op_data, multi_values
            else:
                return lcoe_fitness_scores, multi_fitness_scores, multi_values

        def optQuitClicked(event):
            self.optExit = True
            optDialog.close()

        def chooseClicked(event):
            self.opt_choice = self.sender().text()
            chooseDialog.close()

        def calc_weight(multi_value):
            weight = 0.
            for key, value in self.targets.items():
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
                    if multi_value[key] > value[3]: # high maximum weight
                        w = 1.
                    elif multi_value[key] < value[2]: # low no weight
                        w = 0.
                    else:
                        w = multi_value[key] / (value[3] - value[2])
                weight += w * value[1]
            return weight

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
            fig = plt.figure(title)
            mx = fig.gca(projection='3d')
            plt.title('\n' + title.title() + '\n')
            surf = mx.scatter(data[0], data[1], data[2], picker=1) # enable picking a point
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

        self.optExit = False
        self.setStatus('Optimise processing started')
        details = self.details
        err_msg = ''
        if self.constraints is None:
            try:
                ts = xlrd.open_workbook(self.get_filename(str(self.files[C].text())))
                ws = ts.sheet_by_name(self.sheets[C].currentText())
                self.getConstraints(ws)
                ts.release_resources()
                del ts
            except:
                err_msg = 'Error accessing Constraints'
                self.getConstraints(None)
        if self.generators is None:
            try:
                ts = xlrd.open_workbook(self.get_filename(str(self.files[G].text())))
                ws = ts.sheet_by_name(self.sheets[G].currentText())
                self.getGenerators(ws)
                ts.release_resources()
                del ts
            except:
                if err_msg != '':
                    err_msg += ' and Generators'
                else:
                    err_msg = 'Error accessing Generators'
            self.getGenerators(None)
        if self.optimisation is None:
            try:
                ts = xlrd.open_workbook(self.get_filename(str(self.files[O].text())))
                ws = ts.sheet_by_name(self.sheets[O].currentText())
                self.getoptimisation(ws)
                ts.release_resources()
                del ts
            except:
                if err_msg != '':
                    err_msg += ' and Optimisation'
                else:
                    err_msg = 'Error accessing Optimisation'
                self.getoptimisation(None)
        if err_msg != '':
            self.setStatus(err_msg)
        pm_data_file = self.get_filename(str(self.files[D].text()))
        if pm_data_file[-4:] == '.xls': #xls format
            self.setStatus('Not an option for Powermatch')
            return
        ts = oxl.load_workbook(pm_data_file)
        ws = ts.active
        top_row = ws.max_row - 8760
        if ws.cell(row=top_row, column=1).value != 'Hour' or ws.cell(row=top_row, column=2).value != 'Period':
            self.setStatus('not a Powermatch data spreadsheet')
            self.progressbar.setHidden(True)
            return
        typ_row = top_row - 1
        while typ_row > 0:
            if ws.cell(row=typ_row, column=3).value in tech_names:
                break
            typ_row -= 1
        else:
            self.setStatus('no suitable data')
            return
        icap_row = typ_row + 1
        while icap_row < top_row:
            if ws.cell(row=icap_row, column=1).value == 'Capacity (MW)':
                break
            icap_row += 1
        else:
            self.setStatus('no capacity data')
            return
        optExit = False
        optDialog = QtGui.QDialog()
        grid = QtGui.QGridLayout()
        grid.addWidget(QtGui.QLabel('Adjust load'), 0, 0)
        optLoad = QtGui.QDoubleSpinBox()
        optLoad.setRange(-1, 25)
        optLoad.setDecimals(3)
        optLoad.setSingleStep(.1)
        optLoad.setValue(1)
        rw = 0
        grid.addWidget(optLoad, rw, 1)
        grid.addWidget(QtGui.QLabel('Multiplier for input Load'), rw, 2, 1, 3)
        rw += 1
        grid.addWidget(QtGui.QLabel('Population size'), rw, 0)
        optPopn = QtGui.QSpinBox()
        optPopn.setRange(10, 500)
        optPopn.setSingleStep(10)
        optPopn.setValue(self.optimise_population)
        optPopn.valueChanged.connect(self.changes)
        grid.addWidget(optPopn, rw, 1)
        grid.addWidget(QtGui.QLabel('Size of population'), rw, 2, 1, 3)
        rw += 1
        grid.addWidget(QtGui.QLabel('No. of generations'), rw, 0, 1, 3)
        optGenn = QtGui.QSpinBox()
        optGenn.setRange(10, 500)
        optGenn.setSingleStep(10)
        optGenn.setValue(self.optimise_generations)
        optGenn.valueChanged.connect(self.changes)
        grid.addWidget(optGenn, rw, 1)
        grid.addWidget(QtGui.QLabel('Number of generations (iterations)'), rw, 2, 1, 3)
        rw += 1
        grid.addWidget(QtGui.QLabel('Mutation probability'), rw, 0)
        optMutn = QtGui.QDoubleSpinBox()
        optMutn.setRange(0, 1)
        optMutn.setDecimals(4)
        optMutn.setSingleStep(0.001)
        optMutn.setValue(self.optimise_mutation)
        optMutn.valueChanged.connect(self.changes)
        grid.addWidget(optMutn, rw, 1)
        grid.addWidget(QtGui.QLabel('Add in mutation'), rw, 2, 1, 3)
        rw += 1
        grid.addWidget(QtGui.QLabel('Exit if stable'), rw, 0)
        optStop = QtGui.QSpinBox()
        optStop.setRange(0, 50)
        optStop.setSingleStep(10)
        optStop.setValue(self.optimise_stop)
        optStop.valueChanged.connect(self.changes)
        grid.addWidget(optStop, rw, 1)
        grid.addWidget(QtGui.QLabel('Exit if LCOE/weight remains the same after this many iterations'),
                       rw, 2, 1, 3)
        rw += 1
        grid.addWidget(QtGui.QLabel('Optimisation choice'), rw, 0)
        optCombo = QtGui.QComboBox()
        choices = ['LCOE', 'Multi', 'Both']
        for choice in choices:
            optCombo.addItem(choice)
            if choice == self.optimise_choice:
                optCombo.setCurrentIndex(optCombo.count() - 1)
        grid.addWidget(optCombo, rw, 1)
        grid.addWidget(QtGui.QLabel('Choose type of optimisation'),
                       rw, 2, 1, 3)
        rw += 1
        # for each variable name
        grid.addWidget(QtGui.QLabel('Variable'), rw, 0)
        grid.addWidget(QtGui.QLabel('Weight'), rw, 1)
        grid.addWidget(QtGui.QLabel('Better'), rw, 2)
        grid.addWidget(QtGui.QLabel('Worse'), rw, 3)
        rw += 1
        ndx = grid.count()
        for key in self.targets.keys():
            self.targets[key][4] = ndx
            ndx += 4
        for key, value in self.targets.items():
            if value[2] < 0:
                ud = ' (<html>&uarr;</html>)'
            elif value[3] < 0 or value[3] > value[2]:
                ud = ' (<html>&darr;</html>)'
            else:
                ud = ' (<html>&uarr;</html>)'
            grid.addWidget(QtGui.QLabel(value[0] + ':' + ud), rw, 0)
            weight = QtGui.QDoubleSpinBox()
            weight.setRange(0, 1)
            weight.setDecimals(2)
            weight.setSingleStep(0.05)
            weight.setValue(value[1])
            grid.addWidget(weight, rw, 1)
            if key[-4:] == '_pct':
                minim = QtGui.QDoubleSpinBox()
                minim.setRange(-.1, 1.)
                minim.setDecimals(2)
                minim.setSingleStep(0.1)
                minim.setValue(value[2])
                grid.addWidget(minim, rw, 2)
                maxim = QtGui.QDoubleSpinBox()
                maxim.setRange(-.1, 1.)
                maxim.setDecimals(2)
                maxim.setSingleStep(0.1)
                maxim.setValue(value[3])
                grid.addWidget(maxim, rw, 3)
            else:
                minim = QtGui.QLineEdit()
                minim.setValidator(QtGui.QDoubleValidator())
                minim.validator().setDecimals(2)
                minim.setText(str(value[2]))
                grid.addWidget(minim, rw, 2)
                maxim = QtGui.QLineEdit()
                maxim.setValidator(QtGui.QDoubleValidator())
                maxim.validator().setDecimals(2)
                maxim.setText(str(value[3]))
                grid.addWidget(maxim, rw, 3)
            rw += 1
        quit = QtGui.QPushButton('Quit', self)
        grid.addWidget(quit, rw, 0)
        quit.clicked.connect(optQuitClicked)
        show = QtGui.QPushButton('Proceed', self)
        grid.addWidget(show, rw, 1)
        show.clicked.connect(optDialog.close)
        optDialog.setLayout(grid)
        optDialog.setWindowTitle('Choose Optimisation Parameters')
        optDialog.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        optDialog.exec_()
        if self.optExit: # a fudge to exit
            self.setStatus('Execution aborted.')
            return
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
        re_capacities = []
        re_capacity = 0.
        orig_load = []
        load_col = -1
        orig_tech = {}
        orig_capacity = {}
        dict_order = {} # rely on it being processed in added order
        # first get original renewables generation from data sheet
        for col in range(3, ws.max_column + 1):
            try:
                valu = ws.cell(row=typ_row, column=col).value.replace('-','')
                i = tech_names.index(valu)
            except:
                break # ?? or continue??
            if tech_names[i] != 'Load':
                try:
                    if ws.cell(row=icap_row, column=col).value <= 0:
                        continue
                    orig_capacity[tech_names[i]] = ws.cell(row=icap_row, column=col).value # in case no optimisation
                except:
                    continue
                dict_order[tech_names[i]] = [ws.cell(row=icap_row, column=col).value, 0, 0, 0]
            orig_tech[tech_names[i]] = []
            for row in range(top_row + 1, ws.max_row + 1):
                orig_tech[tech_names[i]].append(ws.cell(row=row, column=col).value)
        ts.close()
        # now add scheduled generation
        for itm in range(self.order.count()):
            tech = str(self.order.item(itm).text())
            if tech not in dict_order.keys():
                dict_order[tech] = [self.generators[tech].capacity, 0, 0, 0]
        capacities = []
        for tech in dict_order.keys():
            dict_order[tech][1] = len(capacities) # first entry
            try:
                if self.optimisation[tech].approach == 'Discrete':
                    capacities.extend(self.optimisation[tech].capacities)
                    dict_order[tech][2] = len(capacities) # last entry
                elif self.optimisation[tech].approach == 'Range':
                    ctr = int((self.optimisation[tech].capacity_max - self.optimisation[tech].capacity_min) / \
                              self.optimisation[tech].capacity_step)
                    if ctr < 1:
                        self.setStatus("Error with Optimisation table entry for '" + tech + "'")
                        return
                    capacities.extend([self.optimisation[tech].capacity_step] * ctr)
                    tot = self.optimisation[tech].capacity_step * ctr + self.optimisation[tech].capacity_min
                    if tot < self.optimisation[tech].capacity_max:
                        capacities.append(self.optimisation[tech].capacity_max - tot)
                    dict_order[tech][2] = len(capacities)
                    dict_order[tech][3] = self.optimisation[tech].capacity_min
                else:
                    dict_order[tech][2] = len(capacities)
            except KeyError as err:
                self.setStatus('Key Error: No Optimisation entry for ' + str(err))
                dict_order[tech] = [orig_capacity[tech], len(capacities), len(capacities) + 5, 0]
                capacities.extend([orig_capacity[tech] / 5.] * 5)
            except:
                err = str(sys.exc_info()[0]) + ',' + str(sys.exc_info()[1])
                self.setStatus('Error: ' + str(err))
                return
        # chromosome = [1] * int(len(capacities) / 2) + [0] * (len(capacities) - int(len(capacities) / 2))
        # we have the original data - from here down we can do our multiple optimisations
        self.adjustby = {'Load': optLoad.value()}
        headers = ['Facility', 'Capacity (MW)', 'Subtotal (MWh)', 'CF', 'Cost ($)', 'LCOE ($/MWh)', 'Emissions (tCO2e)']
        op_data = []
        if self.adjustby['Load'] == 1:
            op_load = orig_tech['Load'][:]
        else:
            op_load = []
            for load in orig_tech['Load']:
                op_load.append(load * self.adjustby['Load'])
        op_load_tot = sum(op_load)
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
        self.opt_progressbar.emit(QtCore.SIGNAL('progress'), 1, 'Processing generation 1')
        population = create_starting_population(population_size, chromosome_length)
        # calculate best score(s) in starting population
        # if do_lcoe best_score = lowest lcoe
        # if do_multi best_multi = lowest weight and if not do_lcoe best_score also = best_weight
        lcoe_scores, multi_scores, multi_values = calculate_fitness(population)
        if do_lcoe:
            best_score = np.min(lcoe_scores)
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
            self.setStatus('Starting Weight: %.2f' % best_multi)
            multi_best_weight = best_multi
            if not do_lcoe:
                best_score = best_multi
        # Add starting best score to progress tracker
        best_score_progress = [best_score]
        best_ctr = 1
        last_score = best_score
        lowest_score = best_score
        ud = '='
        if do_lcoe:
            d_sign = '$'
        else:
            d_sign = ''
        # Now we'll go through the generations of genetic algorithm
        for generation in range(1, maximum_generation):
            tim = (time.time() - start_time)
            if tim < 60:
                tim = ' ( %s %s%.2f ; %.1f secs)' % (ud, d_sign, best_score, tim)
            else:
                tim = ' ( %s %s%.2f ; %.2f mins)' % (ud, d_sign, best_score, tim / 60.)
            self.opt_progressbar.emit(QtCore.SIGNAL('progress'), generation + 1,
                'Processing generation ' + str(generation + 1) + tim)
            QtGui.QApplication.processEvents()
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
            # get back to original size
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
                if multi_best_weight > best_multi:
                    multi_best_weight = best_multi
                if not do_lcoe:
                    best_score = best_multi
            best_score_progress.append(best_score)
            last_score = best_score
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
            if best_score == best_score_progress[-2]:
                ud = '='
            elif best_score < best_score_progress[-2]:
                ud = '<html>&darr;</html>'
            else:
                ud = '<html>&uarr;</html>'
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
        QtGui.QApplication.processEvents()
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
            titl = 'Optimise Muti using Genetic Algorithm'
            ylbl = 'Best Weight'
        if do_multi:
            self.setStatus('Final Weight: %.2f' % multi_best_weight)
        # Plot progress
        x = list(range(1, len(best_score_progress)+ 1))
        rcParams['savefig.directory'] = self.scenarios
        plt.figure(fig)
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
        op_pts = [0, 3, 0, 2, 0, 2, 0]
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
        for gen, value in dict_order.items():
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
            for gen, value in dict_order.items():
                capacity = dict_order[gen][3]
                for c in range(value[1], value[2]):
                    if chromosome[c]:
                        capacity = capacity + capacities[c]
                its[gen].append(capacity / dict_order[gen][0])
        chooseDialog = QtGui.QDialog()
        hbox = QtGui.QHBoxLayout()
        grid = [QtGui.QGridLayout()]
        label = QtGui.QLabel('<b>Facility</b>')
        label.setAlignment(QtCore.Qt.AlignCenter)
        grid[0].addWidget(label, 0, 0)
        for h in range(len(chrom_hdrs)):
            grid.append(QtGui.QGridLayout())
            label = QtGui.QLabel('<b>' + chrom_hdrs[h] + '</b>')
            label.setAlignment(QtCore.Qt.AlignCenter)
            grid[-1].addWidget(label, 0, 0, 1, 2)
        rw = 1
        for key, value in its.items():
            grid[0].addWidget(QtGui.QLabel(key), rw, 0)
            for h in range(len(chrom_hdrs)):
                label = QtGui.QLabel('{:.2f}'.format(value[h]))
                label.setAlignment(QtCore.Qt.AlignRight)
                grid[h + 1].addWidget(label, rw, 0)
                label = QtGui.QLabel('({:,.2f})'.format(value[h] * dict_order[key][0]))
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
            lbl = QtGui.QLabel('<i>' + self.targets[key][0] + '</i>')
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
                label = QtGui.QLabel(txt % amt)
                label.setAlignment(QtCore.Qt.AlignCenter)
                grid[h + 1].addWidget(label, rw, 0, 1, 2)
            rw += 1
        cshow = QtGui.QPushButton('Quit', self)
        grid[0].addWidget(cshow)
        cshow.clicked.connect(chooseDialog.close)
        for h in range(len(chrom_hdrs)):
            button = QtGui.QPushButton(chrom_hdrs[h], self)
            grid[h + 1].addWidget(button, rw, 0, 1, 2)
            button.clicked.connect(chooseClicked) #(chrom_hdrs[h]))
        for gri in grid:
            frame = QtGui.QFrame()
            frame.setFrameStyle(QtGui.QFrame.Box)
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
        return

if "__main__" == __name__:
    app = QtGui.QApplication(sys.argv)
    ex = powerMatch()
    app.exec_()
    app.deleteLater()
    sys.exit()
