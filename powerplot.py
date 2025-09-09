#!/usr/bin/env python3
#
#  Copyright (C) 2019-2025 Sustainable Energy Now Inc., Angus King
#
#  powerplot.py - This file is possibly part of SIREN.
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

import configparser  # decode .ini file
import os
import plotly.graph_objects as go
import plotly
from PyQt5 import QtCore, QtGui, QtWidgets
import sys
from math import ceil, log10, sqrt
import matplotlib
if matplotlib.__version__ > '3.5.1':
    matplotlib.use('Qt5Agg')
else:
    matplotlib.use('TkAgg')
from matplotlib.font_manager import FontProperties
import numpy as np
import matplotlib.pyplot as plt
import displayobject
import displaytable
from colours import Colours, PlotPalette
from credits import fileVersion
from displaytable import Table
from editini import EdtDialog, SaveIni
from getmodels import getModelFile
try:
    from senplot3d import PowerPlot3D
except:
    PowerPlot3D = None
from senutils import ClickableQLabel, getParents, getUser, ListWidget, ssCol, strSplit, techClean, WorkBook
import subprocess
from zoompan import ZoomPanX

mth_labels = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

class MyQDialog(QtWidgets.QDialog):
    ignoreEnter = True

    def keyPressEvent(self, qKeyEvent):
        if qKeyEvent.key() == QtCore.Qt.Key_Return:
            self.ignoreEnter = True
        else:
            self.ignoreEnter = False

    def closeEvent(self, event):
        if self.ignoreEnter:
            self.ignoreEnter = False
            event.ignore()
        else:
            event.accept()


class MyCombo(QtWidgets.QComboBox):
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

    def __init__(self, parent):
        super(MyCombo, self).__init__()
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        ba = event.mimeData().data('application/x-qabstractitemmodeldatalist')
        data_items = self.decode_data(ba)
        for item in data_items:
            if isinstance(item, str):
                for i in range(self.count()):
                    if self.itemText(i) == item:
                        break
                else:
                    self.addItem(item)
        event.acceptProposedAction()


class ChangeFontProp(MyQDialog):
    def __init__(self, what, font=None, legend=None):
        super(ChangeFontProp, self).__init__()
        who = sys.argv[0][sys.argv[0].rfind('/') + 1: sys.argv[0].rfind('.')].title()
        self.what = what.split(' ')[0]
        self._font = font
        self.ignoreEnter = False
         #                        'Weight': [['ultralight', 'light', 'normal', 'regular', 'book', 'medium', 'roman',
         #                                    'semibold', 'demibold', 'demi', 'bold', 'heavy', 'extra bold', 'black'],
         #                                    'normal'],
        self._font_properties = {'Family': [['cursive', 'fantasy', 'monospace', 'sans-serif', 'serif'] ,'sans-serif'],
                                 'Style': [['italic', 'normal', 'oblique'], 'normal'],
                                 'Variant': [['normal', 'small-caps'], 'normal'],
                                 'Stretch': [['ultra-condensed', 'extra-condensed', 'condensed', 'semi-condensed',
                                              'normal', 'semi-expanded', 'expanded', 'extra-expanded',
                                              'ultra-expanded'], 'normal'],
                                 'Weight': [['ultralight', 'light', 'normal', 'bold'], 'normal'],
                                 'Size': [['xx-small', 'x-small', 'small', 'medium', 'large', 'x-large', 'xx-large'],
                                           'small']}
        self._legend_properties = {'legend_ncol': ['Columns', ['1', '2', '3', '4', '5', '6', '7', '8', '9'], '1'],
                                   'legend_on_pie': ['On pie', ['True', 'Pct', 'False'], 'True'],
                                   'legend_side': ['Alignment', ['None', 'Left', 'Right'], 'None']}
#constrained_layout=False
#short_legend
        self._size = [5.79, 6.94, 8.33, 10.0, 12.0, 14.4, 17.28]
        self._legend = None
        self._results = None
        self._color = None
        self._pattern = None
        self.grid = QtWidgets.QGridLayout()
        rw = 0
        if self.what == 'Legend':
            self.legend_items = {}
            self.grid.addWidget(QtWidgets.QLabel('Legend properties'), rw, 0)
            for key, value in self._legend_properties.items():
                self.grid.addWidget(QtWidgets.QLabel(value[0]), rw, 1)
                self.legend_items[key] = QtWidgets.QComboBox()
                for i in range(len(value[1])):
                    self.legend_items[key].addItem(value[1][i])
                    if value[1][i] == value[2]:
                        self.legend_items[key].setCurrentIndex(i)
                self.grid.addWidget(self.legend_items[key], rw, 2)
                rw += 1
            if legend is not None:
                for key in self._legend_properties:
                    try:
                        self.legend_items[key].setCurrentText(legend[key])
                    except:
                        pass
            rw += 1
        else:
            self._color = '#000000' # black
            if font is not None:
                self._font = FontProperties()
                font_str = ''
                for key, value in font.items():
                    if key == 'family':
                        font_str += value.replace('-', "\\-") + ':'
                    elif key == 'color':
                        self._color = value
                    else:
                        font_str += key + '=' + str(value) + ':'
                font_str = font_str[:-1]
            self._font.set_fontconfig_pattern(font_str)
        self.font_items = {}
        self.grid.addWidget(QtWidgets.QLabel('Font properties'), rw, 0)
        self.grid.addWidget(QtWidgets.QLabel('Name'), rw, 1)
        self.name = QtWidgets.QLabel(self._font.get_name())
        self.grid.addWidget(self.name, rw, 2)
        rw += 1
        for key, value in self._font_properties.items():
            self.grid.addWidget(QtWidgets.QLabel(key), rw, 1)
            self.font_items[key] = QtWidgets.QComboBox()
            for i in range(len(value[0])):
                self.font_items[key].addItem(value[0][i])
                if value[0][i] == value[1]:
                    self.font_items[key].setCurrentIndex(i)
            self.grid.addWidget(self.font_items[key], rw, 2)
            rw += 1
        self.font_items['Family'].currentIndexChanged.connect(self.familyChange)
        if self._font is not None:
            try:
                self.font_items['Family'].setCurrentText(self._font.get_family()[0])
                self.font_items['Style'].setCurrentText(self._font.get_style())
                self.font_items['Variant'].setCurrentText(self._font.get_variant())
                self.font_items['Stretch'].setCurrentText(self._font.get_stretch())
                self.font_items['Weight'].setCurrentText(self._font.get_weight())
                sze = self._size.index(round(self._font.get_size(), 2))
                self.font_items['Size'].setCurrentText(self._font_properties['Size'][0][sze])
            except:
                pass
        if self.what != 'Header' and self._color is not None:
            self.grid.addWidget(QtWidgets.QLabel('Colour'), rw, 1)
            self.colbtn = QtWidgets.QPushButton('Color', self)
            self.colbtn.clicked.connect(self.colorChanged)
            if self._color == '#ffffff':
                value = '#808080'
            else:
                value = self._color
            self.colbtn.setStyleSheet('QPushButton {background-color: %s;}' % value)
            self.font_items['Color'] = self.colbtn
            self.grid.addWidget(self.font_items['Color'], rw, 2)
            rw += 1
        quit = QtWidgets.QPushButton('Quit', self)
        self.grid.addWidget(quit, rw, 0)
        quit.clicked.connect(self.quitClicked)
        save = QtWidgets.QPushButton('Save', self)
        self.grid.addWidget(save, rw, 1)
        save.clicked.connect(self.saveClicked)
        reset = QtWidgets.QPushButton('Reset', self)
        self.grid.addWidget(reset, rw, 2)
        reset.clicked.connect(self.resetClicked)
        frame = QtWidgets.QFrame()
        frame.setLayout(self.grid)
        self.scroll = QtWidgets.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtWidgets.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
        self.setWindowTitle('SIREN - ' + who + ' - ' + self.what + ' Properties')
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        QtWidgets.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
        self.show()

    def colorChanged(self):
        col = QtWidgets.QColorDialog.getColor(QtGui.QColor(''))
        if col.isValid():
            self._color = col.name()
            self.colbtn.setStyleSheet('QPushButton {background-color: %s;}' % col.name())
            if self._color.lower() == '#ffffff':
                self.colbtn.setStyleSheet('QPushButton {color: #808080;}')

    def quitClicked(self):
        self.ignoreEnter = False
        self._font = None
        self._legend = None
        self._pattern = None
        self._results = None
        self.close()

    def resetClicked(self):
        for key, value in self._font_properties.items():
            try:
                self.font_items[key].setCurrentText(value[1])
            except:
                pass
        if 'Color' in self.font_items.keys():
            try:
                self.colbtn.setStyleSheet('QPushButton {background-color: #000000;}')
            except:
                pass

    def saveClicked(self):
        self.ignoreEnter = False
        if self.what == 'Legend':
            self._legend = {}
            for key, value in self.legend_items.items():
                self._legend[key] = value.currentText()
            self._results = {}
        self._font = FontProperties()
        self._font.set_family(self.font_items['Family'].currentText())
        self._font.set_style(self.font_items['Style'].currentText())
        self._font.set_variant(self.font_items['Variant'].currentText())
        self._font.set_stretch(self.font_items['Stretch'].currentText())
        self._font.set_weight(self.font_items['Weight'].currentText())
        self._font.set_size(self.font_items['Size'].currentText())
        family = self.font_items['Family'].currentText().replace('-', "\\-")
        self._pattern = f"{family}:" + \
                  f"style={self.font_items['Style'].currentText()}:" + \
                  f"variant={self.font_items['Variant'].currentText()}:" + \
                  f"weight={self.font_items['Weight'].currentText()}:" + \
                  f"stretch={self.font_items['Stretch'].currentText()}:" + \
                  f"size={self._font.get_size()}"
        try:
            self._pattern += f':color={self._color}'
        except:
            pass
        self.close()

    def familyChange(self):
        ffont = FontProperties()
        ffont.set_family(self.font_items['Family'].currentText())
        self.name.setText(ffont.get_name())

    def getValues(self):
        return self._legend, self._results

    def getFont(self):
        return self._font

    def getFontDict(self):
        return self._pattern

class PowerPlot(QtWidgets.QWidget):

    def overlayDialog(self):
        for o in range(len(self.overlays)):
            if self.overlay_button[o].hasFocus():
                break
        col = QtWidgets.QColorDialog.getColor(QtGui.QColor(self.overlay[o]), None, f'Select colour for Overlay{o}')
        if col.isValid():
            self.overlay_colour[o] = col.name()
            self.overlay_button[o].setStyleSheet('QPushButton {background-color: %s; color: %s;}' % (col.name(), col.name()))
            self.updated = True
        return

    def set_colour(self, item, default=None):
        tgt = item.lower()
        alpha = ''
        for alph in self.alpha_word:
            tgt = tgt.replace(alph, '')
        tgt = tgt.replace('  ', ' ')
        tgt = tgt.strip()
        tgt = tgt.lstrip('_')
        if tgt == item.lower():
            if item.lower() in self.colours:
                return self.colours[item.lower()]
            else:
                try:
                    return self.colours[item.lower().replace('_', ' ')]
                except:
                    return default
        else:
            return self.colours[tgt.strip()] + self.alphahex

    def set_fontdict(self, item=None):
        if item is None:
            bits = 'sans-serif:style=normal:variant=normal:weight=normal:stretch=normal:size=10.0:color=black'.split(':')
        else:
            bits = item.split(':')
        fontdict = {}
        for bit in bits:
            bits2 = bit.split('=')
            if len(bits2) == 1:
                bit2 = bits2[0].replace("'", '')
                bit2 = bit2.replace("\\", '')
                fontdict['family'] = bit2
            else:
                try:
                    fontdict[bits2[0]] = float(bits2[1])
                except:
                    fontdict[bits2[0]] = bits2[1]
        return fontdict

    def set_hatch(self, item):
        tgt = item.lower()
        hatch = None
        for alph in self.hatch_word:
            tgt = tgt.replace(alph, '')
        tgt = tgt.replace('  ', ' ')
        tgt = tgt.strip()
        if tgt == item.lower():
            return None
        else:
            # pattern = ['-', '+', 'x', '\\', '|', '/', '*', 'o', 'O', '.']
            return '.'

    def replace_words(self, what, src, tgt):
        words = {'m': ['$MTH$', '$MONTH$'],
                 'y': ['$YEAR$'],
                 's': ['$SHEET$']}
        tgt_str = src
        if tgt == 'find':
            for wrd in words[what]:
                tgt_num = tgt_str.find(wrd)
                if tgt_num >= 0:
                    return tgt_num
                tgt_num = tgt_str.find(wrd.lower())
                if tgt_num >= 0:
                    return tgt_num
            return -1
        else:
            for wrd in words[what]:
                tgt_str = tgt_str.replace(wrd, tgt)
                tgt_str = tgt_str.replace(wrd.lower(), tgt)
        return tgt_str

    def transform_y2(self, y):
        if self.y2_transform is not None:
            return eval(self.y2_transform)
        else:
            return y

    def inverse_y2(self, y):
        if self.y2_inverse is not None:
            return eval(self.y2_inverse)
        elif self.y2_transform is not None:
            inverse = ''
            for c in self.y2_transform:
                if c == '/':
                    inverse += '*'
                elif c == '*':
                    inverse += '/'
                else:
                    inverse += c
            return eval(inverse)
        else:
            return y

    def __init__(self, help='help.html'):
        super(PowerPlot, self).__init__()
        self.help = help
        config = configparser.RawConfigParser()
        if len(sys.argv) > 1:
            self.config_file = getModelFile(sys.argv[1])
        else:
            self.config_file = getModelFile('powerplot.ini')
        config.read(self.config_file)
        parents = []
        self.colours = {}
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
        try:
            colours = config.items('Plot Colors')
            for item, colour in colours:
                itm = item.replace('_', ' ')
                self.colours[itm] = colour
        except:
            pass
        self.restorewindows = False
        try:
            rw = config.get('Windows', 'restorewindows')
            if rw.lower() in ['true', 'yes', 'on']:
                self.restorewindows = True
        except:
            pass
        ifile = ''
        isheet = ''
        columns = []
        self.setup = [False, False]
        self.details = True
        self.book = None
        self.error = False
        # create a colour map based on traffic lights
        cvals = [-2., -1, 2]
        colors = ['red' ,'orange', 'green']
        norm=plt.Normalize(min(cvals), max(cvals))
        tuples = list(zip(map(norm, cvals), colors))
        self.cmap = matplotlib.colors.LinearSegmentedColormap.from_list('', tuples)
        self.cbar = True
        self.cbar2 = True
        self.toprow = None
        self.rows = None
        self.breakdown_row = -1
        self.zone_row = -1
        self.leapyear = False
        iper = '<none>'
        imax = 0
        self.alpha = 0.25
        self.alpha_fill = 1.
        self.alpha_word = ['surplus', 'charge']
        self.margin_of_error = .0001
        self.constrained_layout = False
        self.target = ''
        self.no_of_overlays = 1
        self.overlay = ['<none>']
        self.overlay_colour = ['black']
        self.overlay_type = ['Average']
        self.palette = True
        self.plot_palette = 'random'
        self.legend_on_pie = True
        self.percent_on_pie = True
        self.pie_group = None
        self.legend_side = 'None'
        self.legend_ncol = 1
        self.hatch_word = ['charge']
        self.history = None
        self.max_files = 10
        self.seasons = {}
        self.interval = 24
        self.show_contribution = False
        self.show_correlation = False
        self.select_day = False
        self.plot_12 = False
        self.extras = False
        self.short_legend = ''
        self.legend_font = FontProperties()
        self.legend_font.set_size('small')
        self.fontprops = {}
        self.fontprops['Header'] = self.set_fontdict()
        self.fontprops['Header']['size'] = 16.
        self.fontprops['Label']= self.set_fontdict()
        self.fontprops['Label']['size'] = 10.
        self.fontprops['Ticks']= self.set_fontdict()
        self.fontprops['Ticks']['size'] = 10.
        self.fontprops['Title']= self.set_fontdict()
        self.fontprops['Title']['size'] = 14.
        self.y2_inverse = None
        self.y2_label = ''
        self.y2_transform = None
        ifiles = {}
        try:
            items = config.items('Powerplot')
            for key, value in items:
                if key == 'alpha':
                    try:
                        self.alpha = float(value)
                    except:
                        pass
                elif key == 'alpha_fill':
                    try:
                        self.alpha_fill = float(value)
                    except:
                        pass
                elif key == 'alpha_word':
                    self.alpha_word = value.split(',')
                elif key == 'cbar':
                    if value.lower() in ['false', 'no', 'off']:
                        self.cbar = False
                elif key == 'cbar2':
                    if value.lower() in ['false', 'no', 'off']:
                        self.cbar2 = False
                elif key == 'cmap':
                    if value.find(',') > 0:
                        tuples = list(zip(map(norm, cvals), value.split(',')))
                        self.cmap = matplotlib.colors.LinearSegmentedColormap.from_list('', tuples)
                    else:
                        self.cmap = value
                elif key == 'constrained_layout':
                    if value.lower() in ['true', 'yes', 'on']:
                        self.constrained_layout = True
                elif key == 'cvals':
                    if value.find(',') > 0:
                        cvals = list(map(float, value.split(',')))
                        norm=plt.Normalize(min(cvals), max(cvals))
                        tuples = list(zip(map(norm, cvals), value.split(',')))
                elif key == 'extras':
                    if value.lower() in ['true', 'yes', 'on']:
                        self.extras = True
                elif key == 'file_choices':
                    self.max_files = int(value)
                elif key == 'file_history':
                    self.history = value.split(',')
                elif key[:4] == 'file':
                    ifiles[key[4:]] = value.replace('$USER$', getUser())
                if key == 'hatch_word':
                    self.hatch_word = value.split(',')
                elif key == 'header_font':
                    try:
                        self.fontprops['Header'] = self.set_fontdict(value)
                    except:
                        pass
                elif key == 'margin_of_error':
                    try:
                        self.margin_of_error = float(value)
                    except:
                        pass
                elif key == 'no_of_overlays':
                    try:
                        self.no_of_overlays = int(value)
                        self.overlay = ['<none>'] * self.no_of_overlays
                        self.overlay_colour = ['black'] * self.no_of_overlays
                        self.overlay_type = ['Average'] * self.no_of_overlays
                    except:
                        pass
                elif key == 'label_font':
                    try:
                        self.fontprops['Label'] = self.set_fontdict(value)
                    except:
                        pass
                elif key == 'legend_font':
                    try:
                        self.legend_font.set_fontconfig_pattern(value)
                    except:
                        pass
                elif key == 'legend_on_pie':
                    if value.lower() in ['pct', 'percentage', '%', '%age']:
                        self.legend_on_pie = False
                    elif value.lower() in ['false', 'no', 'off']:
                        self.legend_on_pie = False
                        self.percent_on_pie = False
                elif key == 'legend_ncol':
                    try:
                        self.legend_ncol = int(value)
                    except:
                        pass
                elif key == 'legend_side':
                    try:
                        self.legend_side = value
                    except:
                        pass
                elif key == 'palette':
                    if value.lower() in ['false', 'no', 'off']:
                        self.palette = False
                elif key == 'pie_group':
                    if value.lower() in ['false', 'no', 'off']:
                        self.pie_group = None
                    elif value.lower() in ['true', 'yes', 'on']:
                        self.pie_group = ['bess', 'biomass', 'pv', 'solar', 'wind']
                    else:
                        if value.find(',') > -1:
                            self.pie_group = value.lower.split(',')
                        else:
                            self.pie_group = value.lower.split(' ')
                elif key == 'plot_palette':
                    if value.lower() in ['false', 'no', 'off']:
                        self.plot_palette = ''
                    else:
                        self.plot_palette = value
                elif key == 'plot_12':
                    if value.lower() in ['true', 'yes', 'on']:
                        self.plot_12 = True
                elif key == 'select_day':
                    if value.lower() in ['true', 'yes', 'on']:
                        self.select_day = True
                elif key == 'short_legend':
                    if value.lower() in ['true', 'yes', 'on', '_']:
                        self.short_legend = '_'
                elif key == 'ticks_font':
                    try:
                        self.fontprops['Ticks'] = self.set_fontdict(value)
                    except:
                        pass
                elif key == 'title_font':
                    try:
                        self.fontprops['Title'] = self.set_fontdict(value)
                    except:
                        pass
        except:
            pass
        self.alphahex = hex(int(self.alpha * 255))[2:]
        try:
            items = config.items('Power')
            for item, values in items:
                if item[:6] == 'season':
                    if item == 'season':
                        continue
                    mths = values.split(',')
                    if mths[0] in self.seasons.keys():
                        per = mths[0] + ' Period'
                        self.seasons[per] = self.seasons[mths[0]]
                    self.seasons[mths[0]] = list(map(lambda x: int(x) - 1, mths[1:]))
                elif item[:6] == 'period':
                    if item == 'period':
                        continue
                    mths = values.split(',')
                    if mths[0] in self.seasons.keys():
                        per = mths[0] + ' Period'
                    else:
                        per = mths[0]
                    self.seasons[per] = list(map(lambda x: int(x) - 1, mths[1:]))
        except:
            pass
        try:
            value = config.get('Powermatch', 'show_contribution')
            if value.lower() in ['true', 'yes', 'on']:
                self.show_contribution = True
        except:
            pass
        try:
            value = config.get('Powermatch', 'show_correlation')
            if value.lower() in ['true', 'yes', 'on']:
                self.show_correlation = True
        except:
            pass
        if len(ifiles) > 0:
            if self.history is None:
                self.history = sorted(ifiles.keys(), reverse=True)
            ifile = ifiles[self.history[0]]
        matplotlib.rcParams['savefig.directory'] = os.getcwd()
        self.grid = QtWidgets.QGridLayout()
        self.updated = False
        self.colours_updated = False
        self.log = QtWidgets.QLabel('')
        if isinstance(self.config_file, list):
            self.log.setText('Preferences file: ' + ', '.join(self.config_file))
        else:
            self.log.setText('Preferences file: ' + self.config_file)
        rw = 0
        self.grid.addWidget(QtWidgets.QLabel('Recent Files:'), rw, 0)
        self.files = QtWidgets.QComboBox()
        if ifile != '':
            self.popfileslist(ifile, ifiles)
        self.grid.addWidget(self.files, rw, 1, 1, 5)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('File:'), rw, 0)
        self.file = ClickableQLabel()
        self.file.setStyleSheet("background-color: white; border: 1px inset grey; min-height: 22px; border-radius: 4px;")
        self.file.setText('')
        self.grid.addWidget(self.file, rw, 1, 1, 5)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Sheet:'), rw, 0)
        self.sheet = QtWidgets.QComboBox()
        self.grid.addWidget(self.sheet, rw, 1, 1, 2)
        openfile = QtWidgets.QPushButton('Open File')
        self.grid.addWidget(openfile, rw, 3)
        openfile.clicked.connect(self.openfileClicked)
        listfiles = QtWidgets.QPushButton('List Files')
        self.grid.addWidget(listfiles, rw, 4)
        listfiles.clicked.connect(self.listfilesClicked)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Period:'), rw, 0)
        self.period = QtWidgets.QComboBox()
        self.period.addItem('<none>')
        self.period.addItem('Year')
        for mth in mth_labels:
            self.period.addItem(mth)
        for per in sorted(self.seasons.keys()):
            self.period.addItem(per)
        self.grid.addWidget(self.period, rw, 1, 1, 1)
        self.cperiod = QtWidgets.QComboBox()
        self.cperiod.addItem('<none>')
        for mth in mth_labels:
            self.cperiod.addItem(mth)
        self.grid.addWidget(self.cperiod, rw, 2, 1, 1)
        self.grid.addWidget(QtWidgets.QLabel('(Diurnal profile for Period or Month(s))'), rw, 3, 1, 2)
        if self.select_day:
            rw += 1
            self.grid.addWidget(QtWidgets.QLabel('Day:'), rw, 0)
            self.aday = QtWidgets.QComboBox()
            self.aday.addItem('')
            for d in range(1, 32):
                self.aday.addItem(str(d))
            self.grid.addWidget(self.aday, rw, 1)
            self.adaylbl = QtWidgets.QLabel('')
            self.grid.addWidget(self.adaylbl, rw, 2, 1, 2)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Target:'), rw, 0)
        self.targets = QtWidgets.QComboBox()
        self.grid.addWidget(self.targets, rw, 1, 1, 2)
        self.grid.addWidget(QtWidgets.QLabel('(e.g. Load)    Width: / Style'), rw, 3)
        self.tgtSpin = QtWidgets.QDoubleSpinBox()
        self.tgtSpin.setDecimals(1)
        self.tgtSpin.setRange(1., 4.)
        self.tgtSpin.setSingleStep(.5)
        self.tgtSpin.setValue(2.5)
        self.grid.addWidget(self.tgtSpin, rw, 4)
        self.tgtLine = QtWidgets.QComboBox()
        self.tgtLine.addItem('solid')
        self.tgtLine.addItem('dashed')
        self.tgtLine.addItem('dotted')
        self.tgtLine.addItem('dashdot')
        self.tgtLine.setCurrentIndex(0)
        self.grid.addWidget(self.tgtLine, rw, 5)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Overlays:'), rw, 0)
        self.overlays = []
        self.overlay_button = []
        self.ovrSpin = []
        self.ovrLine = []
        for o in range(self.no_of_overlays):
            self.overlays.append(MyCombo(self))
            self.overlays[o].addItem('<none>')
            self.overlays[o].addItem('Charge')
            self.overlays[o].addItem('Underlying Load')
            self.overlays[o].setCurrentIndex(0)
            self.grid.addWidget(self.overlays[o], rw, 1, 1, 1)
            self.grid.addWidget(QtWidgets.QLabel('Colour / Width / Style:'), rw, 2)
            self.overlay_button.append(QtWidgets.QPushButton(f'Overlay{o}', self))
            self.overlay_button[o].setStyleSheet('QPushButton {background-color: %s; color: %s;}' % (self.overlay_colour[o], self.overlay_colour[o]))
            self.overlay_button[o].clicked.connect(self.overlayDialog)
            self.grid.addWidget(self.overlay_button[o], rw, 3)
            self.ovrSpin.append(QtWidgets.QDoubleSpinBox())
            self.ovrSpin[o].setDecimals(1)
            self.ovrSpin[o].setRange(1., 4.)
            self.ovrSpin[o].setSingleStep(.5)
            self.ovrSpin[o].setValue(1.5)
            self.grid.addWidget(self.ovrSpin[o], rw, 4)
            self.ovrLine.append(QtWidgets.QComboBox())
            self.ovrLine[o].addItem('solid')
            self.ovrLine[o].addItem('dashed')
            self.ovrLine[o].addItem('dotted')
            self.ovrLine[o].addItem('dashdot')
            self.ovrLine[o].setCurrentIndex(2)
            self.grid.addWidget(self.ovrLine[o], rw, 5)
            rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Header:'), rw, 0)
        self.suptitle = QtWidgets.QLineEdit('')
        self.grid.addWidget(self.suptitle, rw, 1, 1, 2)
        hdrfClicked = QtWidgets.QPushButton('Header Font', self)
        self.grid.addWidget(hdrfClicked, rw, 3)
        hdrfClicked.clicked.connect(self.doFont)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Title:'), rw, 0)
        self.title = QtWidgets.QLineEdit('')
        self.grid.addWidget(self.title, rw, 1, 1, 2)
        ttlfClicked = QtWidgets.QPushButton('Title Font', self)
        self.grid.addWidget(ttlfClicked, rw, 3)
        ttlfClicked.clicked.connect(self.doFont)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Maximum:'), rw, 0)
        self.maxSpin = QtWidgets.QDoubleSpinBox()
        self.maxSpin.setDecimals(2)
        self.maxSpin.setRange(0, 10000)
        self.maxSpin.setSingleStep(500)
        self.grid.addWidget(self.maxSpin, rw, 1)
        self.grid.addWidget(QtWidgets.QLabel('(Handy if you want to produce a series of plots)'), rw, 2, 1, 3)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Type of Plot:'), rw, 0)
        plots = ['Cumulative', 'Bar Chart', 'Heat Map', 'Line Chart', 'Load Duration', 'Pie Chart', 'Step Chart', 'Step Day Chart',
                 'Step Line Chart']
        if PowerPlot3D is not None:
            plots.append('3D Surface Chart')
        self.plottype = QtWidgets.QComboBox()
        for plot in plots:
             self.plottype.addItem(plot)
        self.plottype.setCurrentIndex(2) # default to cumulative
        self.grid.addWidget(self.plottype, rw, 1) #, 1, 2)
        self.grid.addWidget(QtWidgets.QLabel('(Stacked except for Line Charts)'), rw, 2, 1, 2)
        self.grid.addWidget(QtWidgets.QLabel(' Line Charts Width / Style:'), rw, 3, 1, 2)
        self.linSpin = QtWidgets.QDoubleSpinBox()
        self.linSpin.setDecimals(1)
        self.linSpin.setRange(1., 4.)
        self.linSpin.setSingleStep(.5)
        self.linSpin.setValue(1.5)
        self.grid.addWidget(self.linSpin, rw, 4)
        self.linLine = QtWidgets.QComboBox()
        self.linLine.addItem('solid')
        self.linLine.addItem('dashed')
        self.linLine.addItem('dotted')
        self.linLine.addItem('dashdot')
        self.linLine.setCurrentIndex(0)
        self.grid.addWidget(self.linLine, rw, 5)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Percentage:'), rw, 0)
        self.percentage = QtWidgets.QCheckBox()
        self.percentage.setCheckState(QtCore.Qt.Unchecked)
        self.grid.addWidget(self.percentage, rw, 1) #, 1, 2)
        self.grid.addWidget(QtWidgets.QLabel('(Check for percentage distribution)'), rw, 2, 1, 3)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Spill label:'), rw, 0)
        self.spill_label = QtWidgets.QLineEdit('')
        self.grid.addWidget(self.spill_label, rw, 1, 1, 2)
        self.grid.addWidget(QtWidgets.QLabel('(Enter suffix if you want spill labels)'), rw, 3, 1, 3)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Show Grid:'), rw, 0)
        grids = ['Both', 'Horizontal', 'Vertical', 'None']
        self.gridtype = QtWidgets.QComboBox()
        for grid in grids:
             self.gridtype.addItem(grid)
        self.grid.addWidget(self.gridtype, rw, 1) #, 1, 2)
        self.grid.addWidget(QtWidgets.QLabel('(Choose gridlines)'), rw, 2, 1, 3)
        rw += 1
        self.brk_label = QtWidgets.QLabel('Breakdowns:')
        self.grid.addWidget(self.brk_label, rw, 0)
        brk_h = self.spill_label.sizeHint().height() * 2
        self.brk_order = ListWidget(self)
        self.brk_order.setFixedHeight(brk_h)
        self.grid.addWidget(self.brk_order, rw, 1, 1, 2)
        self.brk_ignore = ListWidget(self)
        self.brk_ignore.setFixedHeight(brk_h)
        self.grid.addWidget(self.brk_ignore, rw, 3, 1, 2)
        self.brk_label.setHidden(True)
        self.brk_order.setHidden(True)
        self.brk_ignore.setHidden(True)
        rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Column Order:\n(move to right\nto exclude)'), rw, 0)
        self.order = ListWidget(self)
        self.grid.addWidget(self.order, rw, 1, 1, 2)
        self.ignore = ListWidget(self)
        self.grid.addWidget(self.ignore, rw, 3, 1, 2)
        self.grid.addWidget(QtWidgets.QLabel(' '), rw, 5)
        if ifile != '':
            self.get_file_config(self.history[0])
        self.files.currentIndexChanged.connect(self.filesChanged)
        self.file.clicked.connect(self.fileChanged)
        self.period.currentIndexChanged.connect(self.periodChanged)
        self.cperiod.currentIndexChanged.connect(self.somethingChanged)
        self.files.currentIndexChanged.connect(self.targetChanged)
        self.sheet.currentIndexChanged.connect(self.sheetChanged)
        self.targets.currentIndexChanged.connect(self.targetChanged)
        self.tgtSpin.valueChanged.connect(self.somethingChanged)
        self.tgtLine.currentIndexChanged.connect(self.somethingChanged)
        for o in range(self.no_of_overlays):
            self.overlays[o].currentIndexChanged.connect(self.overlayChanged)
            self.ovrSpin[o].valueChanged.connect(self.somethingChanged)
            self.ovrLine[o].currentIndexChanged.connect(self.somethingChanged)
        self.linSpin.valueChanged.connect(self.somethingChanged)
        self.linLine.currentIndexChanged.connect(self.somethingChanged)
        self.suptitle.textChanged.connect(self.somethingChanged)
        self.title.textChanged.connect(self.somethingChanged)
        self.maxSpin.valueChanged.connect(self.somethingChanged)
        self.plottype.currentIndexChanged.connect(self.somethingChanged)
        self.gridtype.currentIndexChanged.connect(self.somethingChanged)
        self.percentage.stateChanged.connect(self.somethingChanged)
        self.spill_label.textChanged.connect(self.somethingChanged)
        self.order.itemSelectionChanged.connect(self.somethingChanged)
        self.brk_order.itemSelectionChanged.connect(self.somethingChanged)
        rw += 1
        lblfClicked = QtWidgets.QPushButton('Label Font', self)
        self.grid.addWidget(lblfClicked, rw, 0)
        lblfClicked.clicked.connect(self.doFont)
        ticfClicked = QtWidgets.QPushButton('Ticks Font', self)
        self.grid.addWidget(ticfClicked, rw, 1)
        ticfClicked.clicked.connect(self.doFont)
        msg_palette = QtGui.QPalette()
        msg_palette.setColor(QtGui.QPalette.Foreground, QtCore.Qt.red)
        self.log.setPalette(msg_palette)
        self.grid.addWidget(self.log, rw, 2, 1, 4)
        rw += 1
        quit = QtWidgets.QPushButton('Done', self)
        self.grid.addWidget(quit, rw, 0)
        quit.clicked.connect(self.quitClicked)
        QtWidgets.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
        pp = QtWidgets.QPushButton('Plot', self)
        self.grid.addWidget(pp, rw, 1)
        pp.clicked.connect(self.ppClicked)
        QtWidgets.QShortcut(QtGui.QKeySequence('p'), self, self.ppClicked)
        cb = QtWidgets.QPushButton('Colours', self)
        self.grid.addWidget(cb, rw, 2)
        cb.clicked.connect(self.editColours)
        legendClicked = QtWidgets.QPushButton('Legend Properties', self)
        self.grid.addWidget(legendClicked, rw, 3)
        legendClicked.clicked.connect(self.doFont)
        editini = QtWidgets.QPushButton('Preferences', self)
        self.grid.addWidget(editini, rw, 4)
        editini.clicked.connect(self.editIniFile)
        help = QtWidgets.QPushButton('Help', self)
        self.grid.addWidget(help, rw, 5)
        help.clicked.connect(self.helpClicked)
        QtWidgets.QShortcut(QtGui.QKeySequence('F1'), self, self.helpClicked)
        if self.extras or self.plot_12 or self.show_contribution or self.show_correlation:
            rw += 1
            c = 0
            if self.plot_12:
                c += 1
                pp12 = QtWidgets.QPushButton('Plot 12', self)
                self.grid.addWidget(pp12, rw, c)
                pp12.clicked.connect(self.ppClicked)
            if self.extras:
                c += 1
                extra = QtWidgets.QPushButton('Extras', self)
                self.grid.addWidget(extra, rw, c)
                extra.clicked.connect(self.doExtras)
            if self.show_contribution or self.show_correlation:
                c += 1
                if self.show_contribution:
                    co = 'Contribution'
                else:
                    co = 'Correlation'
                corr = QtWidgets.QPushButton(co, self)
                self.grid.addWidget(corr, rw, c)
                corr.clicked.connect(self.corrClicked)
        frame = QtWidgets.QFrame()
        frame.setLayout(self.grid)
        if self.select_day:
            self.adaylbl.setText('(Diurnal profile for a day of ' + self.period.currentText() + ')')
        self.scroll = QtWidgets.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtWidgets.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
        self.setWindowTitle('SIREN - powerplot (' + fileVersion() + ') - PowerPlot')
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        if self.restorewindows:
            try:
                rw = config.get('Windows', 'powerplot_size').split(',')
                self.resize(int(rw[0]), int(rw[1]))
                mp = config.get('Windows', 'powerplot_pos').split(',')
                self.move(int(mp[0]), int(mp[1]))
            except:
                pass
        else:
            self.center()
            self.resize(int(self.sizeHint().width() * 1.4), int(self.sizeHint().height() * 1.4))
        self.show()

    def center(self):
        frameGm = self.frameGeometry()
        screen = QtWidgets.QApplication.desktop().screenNumber(QtWidgets.QApplication.desktop().cursor().pos())
        centerPoint = QtWidgets.QApplication.desktop().availableGeometry(screen).center()
        frameGm.moveCenter(centerPoint)
        self.move(frameGm.topLeft())

    def get_file_config(self, choice=''):
        ifile = ''
        config = configparser.RawConfigParser()
        config.read(self.config_file)
        try: # get list of files if any
            ifile = config.get('Powerplot', 'file' + choice).replace('$USER$', getUser())
        except:
            pass
        if not os.path.exists(ifile) and not os.path.exists(self.scenarios + ifile):
            if self.book is not None:
                self.book.close()
                self.book = None
            self.log.setText("Can't find file - " + ifile)
            msgbox = QtWidgets.QMessageBox()
            msgbox.setWindowTitle('SIREN - powerplot file not found')
            msgbox.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
            fname = ifile[ifile.rfind('/') + 1:]
            msgbox.setText("Can't find '" + fname + "'\nDo you want to remove it from file history (Y)?")
            msgbox.setDetailedText(self.log.text())
            msgbox.setIcon(QtWidgets.QMessageBox.Question)
            msgbox.setStandardButtons(QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
            reply = msgbox.exec_()
            if reply == QtWidgets.QMessageBox.Yes:
                self.file_cleanup(choice)
                self.files.removeItem(self.files.currentIndex())
            return
        breakdowns = []
        columns = []
        self.period.setCurrentIndex(0)
        self.cperiod.setCurrentIndex(0)
        self.gridtype.setCurrentIndex(0)
        self.plottype.setCurrentIndex(3)
        for o in range(self.no_of_overlays):
            for c in range(self.overlays[o].count() -1, 2, -1):
                self.overlays[o].removeItem(c)
            self.overlays[o].setCurrentIndex(0)
        try: # get list of files if any
            items = config.items('Powerplot')
            for key, value in items:
                if key == 'breakdown' + choice:
                    breakdowns = strSplit(value)
                elif key == 'columns' + choice:
                    columns = strSplit(value)
                elif key == 'cperiod' + choice:
                    i = self.cperiod.findText(value, QtCore.Qt.MatchExactly)
                    if i >= 0 :
                        self.cperiod.setCurrentIndex(i)
                    else:
                        self.cperiod.setCurrentIndex(0)
                elif key == 'file' + choice:
                    ifile = value.replace('$USER$', getUser())
                elif key == 'grid' + choice:
                    self.gridtype.setCurrentIndex(self.gridtype.findText(value))
                elif key == 'line_style' + choice:
                    try:
                        self.linLine.setCurrentIndex(self.linLine.findText(value))
                    except:
                        pass
                elif key == 'line_width' + choice:
                    try:
                        self.linSpin.setValue(float(value))
                    except:
                        pass
                elif key[:7] == 'overlay':
                    for i in range(len(key) -1, -1, -1):
                        if not key[i].isdigit():
                            break
                    if key[i+1:] == choice:
                        bits = key[7:i+1].strip('_').split('_')
                        if bits[-1].isdigit():
                            o = int(bits[-1])
                        else:
                            o = 0
                        if bits[0] == 'colour':
                            self.overlay_colour[o] = value
                            self.overlay_button[o].setStyleSheet('QPushButton {background-color: %s; color: %s;}' % (self.overlay_colour[o], self.overlay_colour[o]))
                        elif bits[0] == 'type':
                            self.overlay_type[o] = value
                        elif bits[0] == 'style':
                            try:
                                self.ovrLine[o].setCurrentIndex(self.ovrLine[o].findText(value))
                            except:
                                pass
                        elif bits[0] == 'width':
                            try:
                                self.ovrSpin[o].setValue(float(value))
                            except:
                                pass
                        else:
                            try:
                                i = self.overlays[o].findText(value, QtCore.Qt.MatchExactly)
                                if i >= 0:
                                    self.overlays[o].setCurrentIndex(i)
                                else:
                                    self.overlays[o].addItem(value)
                                    self.overlays[o].setCurrentIndex(self.overlays[o].count() - 1)
                                self.overlay[o] = value
                            except:
                                pass
                elif key == 'percentage' + choice:
                    if value.lower() in ['true', 'yes', 'on']:
                        self.percentage.setCheckState(QtCore.Qt.Checked)
                    else:
                        self.percentage.setCheckState(QtCore.Qt.Unchecked)
                elif key == 'period' + choice:
                    i = self.period.findText(value, QtCore.Qt.MatchExactly)
                    if i >= 0 :
                        self.period.setCurrentIndex(i)
                    else:
                        self.period.setCurrentIndex(0)
                elif key == 'plot' + choice:
                    if value == 'Linegraph':
                        value = 'Line Chart'
                    elif value == 'Step Plot':
                        value = 'Step Chart'
                    self.plottype.setCurrentIndex(self.plottype.findText(value))
                elif key == 'maximum' + choice:
                    try:
                        self.maxSpin.setValue(int(value))
                    except:
                        self.maxSpin.setValue(0)
                elif key == 'sheet' + choice:
                    isheet = value
                elif key == 'spill_label' + choice:
                    try:
                        self.spill_label.setText(value)
                    except:
                        pass
                elif key == 'target' + choice:
                    try:
                        self.target = value
                    except:
                        pass
                elif key == 'target_style' + choice:
                    try:
                        self.tgtLine.setCurrentIndex(self.tgtLine.findText(value))
                    except:
                        pass
                elif key == 'target_width' + choice:
                    try:
                        self.tgtSpin.setValue(float(value))
                    except:
                        pass
                elif key == 'suptitle' + choice:
                    self.suptitle.setText(value)
                elif key == 'title' + choice:
                    self.title.setText(value)
                elif key == 'y2_label' + choice:
                    self.y2_label = value
                elif key == 'y2_inverse' + choice:
                    self.y2_inverse = value
                elif key == 'y2_transform' + choice:
                    self.y2_transform = value
        except:
             pass
        if ifile != '':
            if self.book is not None:
                self.book.close()
                self.book = None
                self.toprow = None
            self.file.setText(ifile)
            if os.path.exists(ifile):
                self.setSheet(ifile, isheet)
            elif os.path.exists(self.scenarios + ifile):
                self.setSheet(self.scenarios + ifile, isheet)
            else:
                self.log.setText('File not found - ' + ifile)
                # Maybe delete dodgy entry
                self.error = True
                return
            self.setColumns(isheet, columns=columns, breakdowns=breakdowns)
            for column in columns:
                self.check_colour(column, config, add=False)

    def popfileslist(self, ifile, ifiles=None):
        self.setup[1] = True
        if ifiles is None:
             ifiles = {}
             for i in range(self.files.count()):
                 ifiles[self.history[i]] = self.files.itemText(i)
        if self.history is None:
            self.history = ['']
            ifiles = {'': ifile}
        else:
            for i in range(len(self.history) - 1, -1, -1):
                try:
                    if ifile == ifiles[self.history[i]]:
                        ihist = self.history[i]
                        del self.history[i]
                        self.history.insert(0, ihist) # make this entry first
                        break
                except KeyError:
                    del self.history[i]
                except:
                    pass
            else:
            # find new entry
                if len(self.history) >= self.max_files:
                    self.history.insert(0, self.history.pop(-1)) # make last entry first
                else:
                    hist = sorted(self.history)
                    if hist[0] == '':
                        hist = [int(x) for x in hist[1:]]
                        hist.sort()
                        hist.insert(0, '')
                    else:
                        hist = [int(x) for x in hist[1:]]
                        hist.sort()
                    hist = [str(x) for x in hist]
                    if hist[0] != '':
                        ent = ''
                    else:
                        for i in range(1, len(hist)):
                            if str(i) != hist[i]:
                                ent = str(i)
                                break
                        else:
                            ent = str(i + 1)
                    self.history.insert(0, ent)
                ifiles[self.history[0]] = ifile
        self.files.clear()
        for i in range(len(self.history)):
            try:
                self.files.addItem(ifiles[self.history[i]])
            except:
                pass
        self.files.setCurrentIndex(0)
        self.setup[1] = False

    def fileChanged(self):
        if os.path.exists(self.file.text()):
            curfile = self.file.text()
        else:
            curfile = self.scenarios + self.file.text()
        newfile = QtWidgets.QFileDialog.getOpenFileName(self, 'Open file', curfile)[0]
        if newfile != '':
            if self.book is not None:
                self.book.close()
                self.book = None
            self.toprow = None
            isheet = self.sheet.currentText()
            self.setSheet(newfile, isheet)
            if newfile[: len(self.scenarios)] == self.scenarios:
                self.file.setText(newfile[len(self.scenarios):])
            else:
                self.file.setText(newfile)
            self.popfileslist(self.file.text())
            self.setup[1] = False
            self.updated = True

    def filesChanged(self):
        if self.setup[1]:
            return
        self.setup[0] = True
        self.saveConfig()
        self.get_file_config(self.history[self.files.currentIndex()])
        if self.error:
            self.error = False
            return
        self.popfileslist(self.files.currentText())
        self.log.setText("File 'loaded'")
        self.setup[0] = False

    def periodChanged(self):
        if self.period.currentText() not in mth_labels:
            self.cperiod.setCurrentIndex(0)
        if self.select_day:
            self.adaylbl.setText('(Diurnal profile for a day of ' + self.period.currentText() + ')')
        if not self.setup[0]:
            self.updated = True

    def somethingChanged(self):
       # if self.plottype.currentText() == 'Bar Chart' and self.period.currentText() == '<none>' and 1 == 2:
       #     self.plottype.setCurrentIndex(self.plottype.currentIndex() + 1) # set to something else
       # el
        if self.percentage.isChecked() and \
          self.plottype.currentText() in ['Line Chart', 'Step Day Chart', 'Step Line Chart', 'Pie Chart']:
            self.percentage.setCheckState(QtCore.Qt.Unchecked)
        if not self.setup[0]:
            self.updated = True

    def setSheet(self, ifile, isheet):
        if self.book is None:
            try:
                self.book = WorkBook()
                self.book.open_workbook(ifile)
            except:
                self.log.setText("Can't open file - " + ifile)
                return
        ndx = 0
        self.sheet.clear()
        j = -1
        for sht in self.book.sheet_names():
            j += 1
            self.sheet.addItem(sht)
            if sht == isheet:
                ndx = j
        self.sheet.setCurrentIndex(ndx)

    def sheetChanged(self):
        self.toprow = None
        if self.book is None:
            self.book = WorkBook()
            self.book.open_workbook(newfile)
        isheet = self.sheet.currentText()
        if isheet not in self.book.sheet_names():
            self.log.setText("Can't find sheet - " + isheet)
            return
        self.setColumns(isheet)
        self.updated = True

    def overlayChanged(self):
        self.log.setText('')
        for o in range(len(self.overlays)):
            if self.overlays[o].hasFocus():
                break
        overlay = self.overlays[o].currentText()
        if overlay != self.overlay[o]:
            self.overlay[o] = overlay
            self.updated = True

    def targetChanged(self):
        if self.error:
            self.error = False
            return
        target = self.targets.currentText()
        if target != self.target:
            if target == '<none>':
                if self.target != '':
                    self.ignore.addItem(self.target)
                    try:
                        self.ignore.item(self.ignore.count() - 1).setBackground(QtGui.QColor(tgt))
                    except:
                        pass
                self.target = target
            else:
                items = self.order.findItems(target, QtCore.Qt.MatchExactly)
                for item in items:
                    self.order.takeItem(self.order.row(item))
                items = self.ignore.findItems(target, QtCore.Qt.MatchExactly)
                for item in items:
                    self.ignore.takeItem(self.ignore.row(item))
                if self.target != '<none>':
                    self.ignore.addItem(self.target)
                    tgt = self.target.lower()
                    for alph in self.alpha_word:
                        tgt.replace(alph, '')
                    tgt.replace('  ', ' ')
                    try:
                        self.ignore.item(self.ignore.count() - 1).setBackground(QtGui.QColor(self.set_colour(tgt, default='white')))
                    except:
                        pass
                self.target = target
        self.updated = True

    def doFont(self):
        if self.sender().text()[:6] == 'Legend':
            if self.legend_on_pie and self.percent_on_pie:
                lop = 'True'
            elif not self.legend_on_pie and self.percent_on_pie:
                lop = 'Pct'
            else:
                lop = 'False'
            legend_properties = {'legend_ncol': str(self.legend_ncol),
                                 'legend_on_pie': lop,
                                 'legend_side': self.legend_side}
            legend = ChangeFontProp('Legend', self.legend_font, legend_properties)
            legend.exec_()
            values, font_values = legend.getValues()
            if values is None:
                return
            for key, value in values.items():
                if key == 'legend_ncol':
                    self.legend_ncol = int(value)
                elif key == 'legend_on_pie':
                    if value.lower() in ['true', 'yes', 'on']:
                        self.legend_on_pie = True
                        self.percent_on_pie = True
                    if value.lower() in ['pct', 'percentage', '%', '%age']:
                        self.legend_on_pie = False
                        self.percent_on_pie = True
                    elif value.lower() in ['false', 'no', 'off']:
                        self.legend_on_pie = False
                        self.percent_on_pie = False
                else:
                    setattr(self, key, value)
            self.legend_font = legend.getFont()
            self.log.setText('Legend properties updated')
            self.updated = True
        else:
            what = self.sender().text().split(' ')[0]
            font = ChangeFontProp(what, self.fontprops[what])
            font.exec()
            if font.getFontDict() is None:
                return
            self.fontprops[what] = self.set_fontdict(font.getFontDict())
            self.log.setText(what + ' font properties updated')
            self.updated = True
        return

    def doExtras(self):
        extra_dict = {}
        extra_dict['y2_inverse'] = self.y2_inverse
        extra_dict['y2_label'] = self.y2_label
        extra_dict['y2_transform'] = self.y2_transform
        combo = QtWidgets.QComboBox(self)
        otypes = ['Average', 'Sum']
        for h in range(24):
            otypes.append(f'{h:02}:00')
        for o in otypes:
            combo.addItem(o)
        i = combo.findText(self.overlay_type, QtCore.Qt.MatchExactly)
        combo.setCurrentIndex(i)
        extra_dict['overlay_type'] = combo
        dialog = Table(extra_dict, fields=['property', 'value'], title='Extra options', edit=True)
        dialog.exec_()
        values = dialog.getValues()
        if values is None:
            return
        for key, value in values.items():
            if len(value) == 0:
                setattr(self, key, '')
            else:
                try:
                    setattr(self, key, value[0].split('=')[1])
                except:
                    pass
        self.updated = True

    def setColumns(self, isheet, columns=[], breakdowns=[]):
        try:
            ws = self.book.sheet_by_name(isheet)
        except:
            self.log.setText("Can't find sheet - " + isheet)
            return
        tech_row = -1
        row = 0
        self.zone_row = -1
        self.breakdown_row = -1
        ignores = []
        while row < ws.nrows:
            if ws.cell_value(row, 0) in ['Split', 'Breakdown']:
                self.breakdown_row = row
                self.brk_order.clear()
                self.brk_ignore.clear()
                brk_orders = []
                for col in range(2, ws.ncols):
                    if ws.cell_value(row, col) is not None and ws.cell_value(row, col) != '':
                        try:
                            brk_orders.index(ws.cell_value(row, col))
                        except:
                            brk_orders.append(ws.cell_value(row, col))
                for brk_order in brk_orders:
                    if brk_order in breakdowns:
                        self.brk_order.addItem(brk_order)
                    else:
                        self.brk_ignore.addItem(brk_order)
                self.brk_label.setHidden(False)
                self.brk_order.setHidden(False)
                self.brk_ignore.setHidden(False)
            elif ws.cell_value(row, 0) == 'Technology':
                tech_row = row
            elif ws.cell_value(row, 0) == 'Zone':
                self.zone_row = row
            elif ws.cell_value(row, 0) in ['Hour', 'Interval', 'Trading Interval']:
                if ws.cell_value(row, 1) != 'Period':
                    self.log.setText(isheet + ' sheet format incorrect')
                    return
                if ws.cell_value(row, 0) in ['Interval', 'Trading Interval']:
                    self.interval = 48
                else:
                    self.interval = 24
                if tech_row >= 0:
                    self.toprow = [tech_row, row]
                else:
                    self.toprow = [row, row]
                self.rows = ws.nrows - (row + 1)
                oldcolumns = []
                if len(columns) == 0:
                    for col in range(self.order.count()):
                        oldcolumns.append(self.order.item(col).text())
                self.order.clear()
                self.ignore.clear()
                self.targets.clear()
                self.targets.addItem('<none>')
                self.targets.setCurrentIndex(0)
                if tech_row >= 0:
                    the_row = tech_row
                else:
                    the_row = row
                for col in range(ws.ncols -1, 1, -1):
                    try:
                        column = ws.cell_value(the_row, col).replace('\n',' ')
                    except:
                        column = str(ws.cell_value(the_row, col))
                    if self.breakdown_row >= 0:
                        try:
                            i = brk_orders.index(ws.cell_value(self.breakdown_row, col))
                            if i > 0:
                                if self.zone_row < 0:
                                    continue
                                if ws.cell_value(self.zone_row, col) is not None and ws.cell_value(self.zone_row, col) != '':
                                    continue
                        except:
                            pass
                    if self.zone_row > 0 and ws.cell_value(self.zone_row, col) != '' and ws.cell_value(self.zone_row, col) is not None:
                        column = ws.cell_value(self.zone_row, col).replace('\n',' ') + '.' + column
                    if column in oldcolumns:
                        if column not in columns:
                            columns.append(column)
                    if self.targets.findText(column, QtCore.Qt.MatchExactly) >= 0:
                        pass
                    else:
                        self.targets.addItem(column)
                        if column == self.target:
                            itm = self.targets.findText(column, QtCore.Qt.MatchExactly)
                            self.targets.setCurrentIndex(itm)
                    if column in columns:
                        pass
                    else:
                        if column != self.target:
                            if column in ignores:
                                continue
                            ignores.append(column)
                            self.ignore.addItem(column)
                            try:
                                self.ignore.item(self.ignore.count() - \
                                 1).setBackground(QtGui.QColor(self.set_colour(column, default='white')))
                            except:
                                pass
                for column in columns:
                    if column != self.target:
                        self.order.addItem(column)
                        try:
                            self.order.item(self.order.count() - 1).setBackground(QtGui.QColor(self.set_colour(column, default='white')))
                        except:
                            return None
                break
            row += 1
        else:
            self.log.setText(isheet + ' sheet format incorrect')
        if self.breakdown_row < 0:
            self.brk_order.clear()
            self.brk_ignore.clear()
            self.brk_label.setHidden(True)
            self.brk_order.setHidden(True)
            self.brk_ignore.setHidden(True)

    def helpClicked(self):
        dialog = displayobject.AnObject(QtWidgets.QDialog(), self.help,
                 title='Help for powerplot (' + fileVersion() + ')', section='powerplot')
        dialog.exec()

    def openfileClicked(self):
        if os.path.exists(self.file.text()):
            curfile = self.file.text()
        else:
            curfile = self.scenarios + self.file.text()
            if not os.path.exists(curfile):
                self.log.setText(f'File not found - {curfile}')
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            os.startfile(curfile)
        elif sys.platform == 'darwin':
            try:
                subprocess.call('open', curfile)
            except:
                try:
                    subprocess.call('open', '-a', 'Microsoft Excel', curfile)
                except:
                    self.setStatus(f"Can't open {curfile}")
                    return
        elif sys.platform == 'linux2' or sys.platform == 'linux':
            subprocess.call(('xdg-open', curfile))
        self.setStatus(f'{curfile} opened')

    def listfilesClicked(self):
        ifiles = {}
        for i in range(self.files.count()):
             if self.history[i] == '':
                 ifiles[' '] = self.files.itemText(i)
             else:
                 ifiles[self.history[i]] = self.files.itemText(i)
        fields = ['Choice', 'File name']
        dialog = displaytable.Table(ifiles, title='Files', fields=fields, sortby='', edit='delete')
        dialog.exec_()
        if dialog.getValues() is not None:
            ofiles = dialog.getValues()
            dels = []
            for key in ifiles.keys():
                if key not in ofiles.keys():
                    if key == ' ':
                        key = ''
                    dels.append(key)
            for choice in dels:
                self.file_cleanup(choice)
                for i in range(self.files.count()):
                     if self.files.itemText(i) == ifiles[choice]:
                         self.files.removeItem(i)
        del dialog
        self.filesChanged()

    def file_cleanup(self, choice):
        if choice == ' ':
            choice = ''
        line = ''
        for i in range(len(self.history) -1, -1, -1):
            if self.history[i] == choice:
                 del self.history[i]
            else:
                line = self.history[i] + ',' + line
        line = line[:-1]
        lines = ['file_history=' + line]
        for prop in ['breakdown', 'columns', 'cperiod', 'cumulative', 'file', 'grid', 'line_style',
                     'line_width', 'maximum', 'period', 'percentage', 'plot', 'sheet', 'spill_label',
                     'suptitle', 'target', 'target_style', 'target_width', 'title', 'y2_label',
                     'y2_inverse', 'y2_transform' ]:
            lines.append(prop + choice + '=')
        for o in range(self.no_of_overlays):
            if o == 0:
                ovr = ''
            else:
                ovr = f'_{o}_'
            for prop in ['', '_colour', '_style', '_type', '_width']:
                lines.append('overlay' + prop + ovr + choice + '=')
        updates = {'Powerplot': lines}
        SaveIni(updates, ini_file=self.config_file)
        self.log.setText('File removed from file history')

    def quitClicked(self):
        if self.book is not None:
            self.book.close()
        if not self.updated and not self.colours_updated:
            self.close()
        self.saveConfig()
        if self.restorewindows:
            updates = {}
            lines = []
            add = int((self.frameSize().width() - self.size().width()) / 2)   # need to account for border
            lines.append('powerplot_pos=%s,%s' % (str(self.pos().x() + add), str(self.pos().y() + add)))
            lines.append('powerplot_size=%s,%s' % (str(self.width()), str(self.height())))
            updates = {'Windows': lines}
            SaveIni(updates, ini_file=self.config_file)
        self.close()

    def closeEvent(self, event):
        try:
            plt.close('all')
        except:
            pass
        event.accept()

    def saveConfig(self):
        updates = {}
        if self.updated:
            config = configparser.RawConfigParser()
            config.read(self.config_file)
            choice = self.history[0]
            save_file = self.file.text().replace(getUser(), '$USER$')
            try:
                self.max_files = int(config.get('Powerplot', 'file_choices'))
            except:
                pass
            lines = []
            lines.append('legend_font=' + self.legend_font.get_fontconfig_pattern())
            lines.append('legend_on_pie=')
            if not self.legend_on_pie and self.percent_on_pie:
                lines[-1] += 'Pct'
            elif not self.legend_on_pie and not self.percent_on_pie:
                lines[-1] += 'False'
            lines.append('legend_ncol=')
            if self.legend_ncol != 1:
                lines[-1] += str(self.legend_ncol)
            lines.append('legend_side=')
            if self.legend_side != 'None':
                lines[-1] += self.legend_side
            for key, value in self.fontprops.items():
                lines.append(key.lower() + '_font=')
                for key2, value2 in value.items():
                    lines[-1] += f'{key2}={value2}:'
                lines[-1] = lines[-1][:-1]
            if len(self.history) > 0:
                line = ''
                for itm in self.history:
                    line += itm + ','
                line = line[:-1]
                lines.append('file_history=' + line)
            cols = 'breakdown' + choice + '='
            for col in range(self.brk_order.count()):
                try:
                    if self.brk_order.item(col).text().index(',') >= 0:
                        try:
                            if self.brk_order.item(col).text().index("'") >= 0:
                                qte = '"'
                        except:
                            qte = "'"
                except:
                    qte = ''
                cols += qte + self.brk_order.item(col).text() + qte + ','
            if cols[-1] != '=':
                cols = cols[:-1]
            lines.append(cols)
            cols = 'columns' + choice + '='
            for col in range(self.order.count()):
                try:
                    if self.order.item(col).text().index(',') >= 0:
                        try:
                            if self.order.item(col).text().index("'") >= 0:
                                qte = '"'
                        except:
                            qte = "'"
                except:
                    qte = ''
                cols += qte + self.order.item(col).text() + qte + ','
            if cols[-1] != '=':
                cols = cols[:-1]
            lines.append(cols)
            lines.append('cperiod' + choice + '=')
            if self.cperiod.currentText() != '<none>':
                lines[-1] = lines[-1] + self.cperiod.currentText()
            lines.append('file' + choice + '=' + self.file.text().replace(getUser(), '$USER$'))
            lines.append('grid' + choice + '=' + self.gridtype.currentText())
            lines.append('line_style' + choice + '=')
            if self.linLine.currentText() != 'dotted':
                lines[-1] = lines[-1] + self.linLine.currentText()
            lines.append('line_width' + choice + '=')
            if self.linSpin.value() != 1.5:
                lines[-1] = lines[-1] + str(self.linSpin.value())
            lines.append('maximum' + choice + '=')
            if self.maxSpin.value() != 0:
                lines[-1] = lines[-1] + str(self.maxSpin.value())
            for o in range(self.no_of_overlays):
                if o == 0:
                    ovr = ''
                else:
                    ovr = f'_{o}_'
                if self.overlay[o] == '<none>':
                    overlay = ''
                else:
                    overlay = self.overlay[o]
                lines.append('overlay' + ovr + choice + '=' + overlay)
                if self.overlay_colour[o] == 'black':
                    overlay_colour = ''
                else:
                    overlay_colour = self.overlay_colour[o]
                lines.append('overlay_colour' + ovr + choice + '=' + overlay_colour)
                lines.append('overlay_style' + ovr + choice + '=')
                if self.ovrLine[o].currentText() != 'dotted':
                    lines[-1] = lines[-1] + self.ovrLine[o].currentText()
      #          lines.append('overlay_type' + ovr + choice + '=' + '' if self.overlay_type[o] == 'Average' else self.overlay_type[o])
                if self.overlay_type[o] == 'Average':
                    overlay_type = ''
                else:
                    overlay_type = self.overlay_type[o]
                lines.append('overlay_type' + ovr + choice + '=' + overlay_type)
                lines.append('overlay_width' + ovr + choice + '=')
                if self.ovrSpin[o].value() != 1.5:
                    lines[-1] = lines[-1] + str(self.ovrSpin[o].value())
            lines.append('percentage' + choice + '=')
            if self.percentage.isChecked():
                lines[-1] = lines[-1] + 'True'
            lines.append('period' + choice + '=')
            if self.period.currentText() != '<none>':
                lines[-1] = lines[-1] + self.period.currentText()
            lines.append('plot' + choice + '=' + self.plottype.currentText())
            lines.append('sheet' + choice + '=' + self.sheet.currentText())
            lines.append('spill_label' + choice + '=' + self.spill_label.text())
            lines.append('suptitle' + choice + '=' + self.suptitle.text())
            lines.append('target' + choice + '=' + self.target)
            lines.append('target_style' + choice + '=')
            if self.tgtLine.currentText() != 'solid':
                lines[-1] = lines[-1] + self.tgtLine.currentText()
            lines.append('target_width' + choice + '=')
            if self.tgtSpin.value() != 2.5:
                lines[-1] = lines[-1] + str(self.tgtSpin.value())
            lines.append('title' + choice + '=' + self.title.text())
            try:
                lines.append('y2_label' + choice + '=' + self.y2_label)
            except:
                pass
            try:
                lines.append('y2_inverse' + choice + '=' + self.y2_inverse)
            except:
                pass
            try:
                lines.append('y2_transform' + choice + '=' + self.y2_transform)
            except:
                pass
            updates['Powerplot'] = lines
        if self.colours_updated:
            lines = []
            for key, value in self.colours.items():
                if value != '':
                    lines.append(key.replace(' ', '_') + '=' + value)
            updates['Plot Colors'] = lines
        SaveIni(updates, ini_file=self.config_file)
        self.updated = False
        self.colours_updated = False

    def editColours(self, color=False):
        # if they've selected some items I'll create a palette of colours for them
        palette = []
        if self.palette:
            for item in self.order.selectedItems():
                palette.append(item.text())
            for item in self.ignore.selectedItems():
                palette.append(item.text())
        dialr = Colours(section='Plot Colors', ini_file=self.config_file, add_colour=color,
                        palette=palette)
        if not dialr.cancelled:
            dialr.exec_()
        self.colours = {}
        config = configparser.RawConfigParser()
        config.read(self.config_file)
        try:
            colours = config.items('Plot Colors')
            for item, colour in colours:
                itm = item.replace('_', ' ')
                self.colours[itm] = colour
        except:
            pass
        for c in range(self.order.count()):
            col = self.order.item(c).text()
            try:
                self.order.item(c).setBackground(QtGui.QColor(self.set_colour(col, default='white')))
            except:
                pass
        for c in range(self.ignore.count()):
            col = self.ignore.item(c).text()
            try:
                self.ignore.item(c).setBackground(QtGui.QColor(self.set_colour(col, default='white')))
            except:
                pass

    def check_colour(self, colour, config, add=True):
        colr = colour.lower()
        if colr in self.colours.keys():
            return True
        try:
            colr2 = self.set_colour(colr)
            return True
        except:
            pass
        colr2 = colr.replace('_', ' ')
        if colr2 in self.colours.keys():
            return True
        elif self.plot_palette[0].lower() == 'r':
            # new approach to generate random colour if not in [Plot Colors]
            r = lambda: random.randint(0,255)
            new_colr = '#%02X%02X%02X' % (r(),r(),r())
            self.colours[colr] = new_colr
            return True
        # previous approach may ask for new colours to be stored in .ini file
        if config is not None:
            try:
                amap = config.get('Map', 'map_choice')
                tgt_colr = config.get('Colors' + amap, colr2)
                self.colours[colr] = tgt_colr
                self.colours_updated = True
                return True
            except:
                pass
            try:
                tgt_colr = config.get('Colors', colr2)
                self.colours[colr] = tgt_colr
                self.colours_updated = True
                return True
            except:
                pass
            if not add:
                return False
            self.editColours(color=colour)
            if colr not in self.colours.keys():
                self.log.setText('No colour for ' + colour)
                return False
            return True
        if not add:
            return False
        self.editColours(color=colour)
        if colr not in self.colours.keys():
            self.log.setText('No colour for ' + colour)
            return False
        return True

    def editIniFile(self):
        if isinstance(self.config_file, list):
            config_file = self.config_file[-1]
        else:
            config_file = self.config_file
        before = os.stat(config_file).st_mtime
        dialr = EdtDialog(self.config_file, section='[Powerplot]')
        dialr.exec_()
        after = os.stat(config_file).st_mtime
        if after == before:
            return
        config = configparser.RawConfigParser()
        config.read(self.config_file)
        try:
            self.alpha = 0.25
            self.alpha_fill = 1.
            self.constrained_layout = False
            self.margin_of_error = .0001
            items = config.items('Powerplot')
            for key, value in items:
                if key == 'alpha':
                    try:
                        self.alpha = float(value)
                    except:
                        pass
                elif key == 'alpha_fill':
                    try:
                        self.alpha_fill = float(value)
                    except:
                        pass
                elif key == 'alpha_word':
                    self.alpha_word = value.split(',')
                elif key == 'cbar':
                    if value.lower() in ['false', 'no', 'off']:
                        self.cbar = False
                    else:
                        self.cbar = True
                elif key == 'cbar2':
                    if value.lower() in ['false', 'no', 'off']:
                        self.cbar2 = False
                    else:
                        self.cbar2 = True
                elif key == 'cmap':
                    self.cmap = value
                elif key == 'constrained_layout':
                    if value.lower() in ['true', 'yes', 'on']:
                        self.constrained_layout = True
                    else:
                        pass
                elif key == 'hatch_word':
                    self.hatch_word = value.split(',')
                elif key == 'header_font':
                    try:
                        self.fontprops['Header'] = self.set_fontdict(value)
                    except:
                        pass
                elif key == 'label_font':
                    try:
                        self.fontprops['Label'] = self.set_fontdict(value)
                    except:
                        pass
                elif key == 'margin_of_error':
                    try:
                        self.margin_of_error = float(value)
                    except:
                        pass
                elif key == 'legend_font':
                    try:
                        self.legend_font.set_fontconfig_pattern(value)
                    except:
                        pass
                elif key == 'legend_on_pie':
                    if value.lower() in ['pct', 'percentage', '%', '%age']:
                        self.legend_on_pie = False
                        self.percent_on_pie = True
                    elif value.lower() in ['false', 'no', 'off']:
                        self.legend_on_pie = False
                        self.percent_on_pie = False
                    else:
                        self.legend_on_pie = True
                        self.percent_on_pie = True
                elif key == 'legend_ncol':
                    try:
                        self.legend_ncol = int(value)
                    except:
                        pass
                elif key == 'palette':
                    if value.lower() in ['false', 'no', 'off']:
                        self.palette = False
                    else:
                        self.palette = True
                elif key == 'short_legend':
                    if value.lower() in ['true', 'yes', 'on', '_']:
                        self.short_legend = '_'
                    elif value.lower() in ['false', 'no', 'off']:
                        self.short_legend = ''
                elif key == 'ticks_font':
                    try:
                        self.fontprops['Ticks'] = self.set_fontdict(value)
                    except:
                        pass
                elif key == 'title_font':
                    try:
                        self.fontprops['Title'] = self.set_fontdict(value)
                    except:
                        pass
        except:
            pass
        self.log.setText(config_file + ' edited. Reload may be required.')

    def ppClicked(self):
        def set_ylimits(ymin, ymax):
            if self.maxSpin.value() > 0:
                ymax = self.maxSpin.value()
            else:
                try:
                    rndup = pow(10, round(log10(ymax * 1.5) - 1)) / 2
                    ymax = ceil(ymax / rndup) * rndup
                except:
                    pass
            if ymin < 0:
                if self.percentage.isChecked():
                    ymin = 0
                else:
                    try:
                        rndup = pow(10, round(log10(abs(ymin) * 1.5) - 1)) / 2
                        ymin = -ceil(abs(ymin) / rndup) * rndup
                    except:
                        pass
            return [ymin, ymax]

        self.log.setText('')
        QtCore.QCoreApplication.processEvents()
        if self.book is None:
            self.log.setText('Error accessing Workbook.')
            return
        if self.order.count() == 0 and self.target == '<none>' and self.plottype.currentText() != 'Line Chart' \
          and self.plottype.currentText() != 'Step Line Chart':
            self.log.setText('Nothing to plot.')
            return
        isheet = self.sheet.currentText()
        if isheet == '':
            self.log.setText('Sheet not set.')
            return
        config = configparser.RawConfigParser()
        config.read(self.config_file)
        for c in range(self.order.count()):
            if not self.check_colour(self.order.item(c).text(), config):
                return
        if self.target != '<none>':
            if not self.check_colour(self.target, config):
                return
            if not self.check_colour('shortfall', config):
                return
        if self.plottype.currentText() == '3D Surface Chart':
            d3_contours = False
            d3_background = True
            d3_html = False
            d3_months = 3
            d3_aspectmode = 'auto'
            try:
                d3_aspectmode = config.get('Powerplot', 'aspectmode_3d').lower()
            except:
                pass
            try:
                chk = config.get('Powerplot', 'background_3d')
                if chk.lower() in ['false', 'off', 'no']:
                    d3_background = False
            except:
                pass
            try:
                chk = config.get('Powerplot', 'contours_3d')
                if chk.lower() in ['true', 'on', 'yes']:
                    d3_contours = True
            except:
                pass
            try:
                chk = config.get('Powerplot', 'html_3d')
                if chk.lower() in ['true', 'on', 'yes']:
                    d3_html = True
            except:
                pass
            try:
                d3_months = int(config.get('Powerplot', 'months_3d'))
            except:
                pass
        del config
        do_12 = False
        do_12_labels = None
        try:
            if self.sender().text() == 'Plot 12':
                do_12 = True
                do_12_save = matplotlib.rcParams['figure.figsize']
                do_12_labels = mth_labels[:]
        except:
            pass
        i = self.file.text().rfind('/')
        if i > 0:
            matplotlib.rcParams['savefig.directory'] = self.file.text()[:i + 1]
        else:
            matplotlib.rcParams['savefig.directory'] = self.scenarios
        if self.gridtype.currentText() == 'Both':
            gridtype = 'both'
        elif self.gridtype.currentText() == 'Horizontal':
            gridtype = 'y'
        elif self.gridtype.currentText() == 'Vertical':
            gridtype = 'x'
        else:
            gridtype = ''
        ws = self.book.sheet_by_name(isheet)
        if self.toprow is None:
            tech_row = -1
            row = 0
            self.zone_row =  -1
            self.breakdown_row = -1
            while row < ws.nrows:
                if ws.cell_value(row, 0) in ['Split', 'Breakdown']:
                    self.breakdown_row = row
                elif ws.cell_value(row, 0) == 'Technology':
                    tech_row = row
                elif ws.cell_value(row, 0) == 'Zone':
                    self.zone_row = row
                elif ws.cell_value(row, 0) in ['Hour', 'Interval', 'Trading Interval']:
                    if ws.cell_value(row, 1) != 'Period':
                        self.log.setText(isheet + ' sheet format incorrect')
                        return
                    if tech_row >= 0:
                        self.toprow = [tech_row, row]
                    else:
                        self.toprow = [row, row]
                    self.rows = ws.nrows - (row + 1)
                    break
                row += 1
            if self.toprow is None:
                self.log.setText(isheet + ' sheet format incorrect')
                return
        ignore_end = True
        if ignore_end:
            for row in range(ws.nrows -1, -1, -1):
                if ws.cell_value(row, 0) is not None:
                    self.rows = row
                    break
        try:
            year = int(ws.cell_value(self.toprow[1] + 1, 1)[:4])
            if year % 4 == 0 and year % 100 != 0 or year % 400 == 0:
                self.leapyear = True
            else:
                self.leapyear = False
        except:
            self.leapyear = False
            year = ''
        the_days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
        if self.leapyear: #rows == 8784: # leap year
            feb29_row = self.toprow[1] + (the_days[0] + the_days[1]) * self.interval + 1
            try:
                if ws.cell_value(feb29_row, 1)[5:10] == '02-29':
                    the_days[1] = 29
            except:
                pass
  #          if ws.nrows - self.toprow[1] - 1 > 8760:
   #             the_days[1] = 29
        hr_labels = ['0:00', '4:00', '8:00', '12:00', '16:00', '20:00', '23:00']
        nocolour = []
        for o in range(self.order.count()):
            if self.order.item(o).text().lower() not in self.colours.keys():
                nocolour.append(self.order.item(o).text().lower())
        if len(nocolour) > 0 and self.plot_palette != '':
            more_colour = PlotPalette(nocolour, palette=self.plot_palette)
            self.colours = {**self.colours, **more_colour}
        if self.plottype.currentText() == '3D Surface Chart':
            order = []
            for o in range(self.order.count()):
                order.append(self.order.item(o).text())
            if d3_html:
                saveit = self.file.text()[:self.file.text().rfind('.')] + '.html'
            else:
                saveit = None
            colours = {}
            nocolour = []
            for o in range(len(order)):
                try:
                    colours[order[o].lower()] = self.colours[order[o].lower()]
                except:
                    nocolour.append(order[o].lower().lower())
            if len(nocolour) > 0:
                if len(nocolour) > 0:
                    more_colour = PlotPalette(nocolour, palette=self.plot_palette)
                    colours = {**colours, **more_colour}
            PowerPlot3D(colours, self.cperiod.currentText(), self.interval, order, self.period.currentText(), self.rows, self.seasons,
                        the_days, self.title.text(), self.toprow, ws, year, self.zone_row,
                        html=saveit, background=d3_background, contours=d3_contours, months=d3_months, aspectmode=d3_aspectmode)
            self.log.setText("3D Chart complete (You'll need to close the browser window yourself).")
            return
        figname = self.plottype.currentText().lower().replace(' ','') + '_' + str(year)
        breakdowns = []
        suptitle = self.suptitle.text()
        if self.breakdown_row >= 0:
            for c in range(self.brk_order.count()):
                breakdowns.append(self.brk_order.item(c).text())
        label = []
        try:
            self.header_color = self.fontprops['Header']['color']
            del self.fontprops['Header']['color']
        except:
            pass
        if self.period.currentText() == '<none>' \
          or (self.period.currentText() == 'Year' and self.plottype.currentText() == 'Heat Map') \
          or self.plottype.currentText() == 'Step Day Chart': # full year of hourly figures
            m = 0
            d = 1
            day_labels = []
            flex_on = True
            if self.plottype.currentText() == 'Step Day Chart':
                days_per_label = 7
                flex_on = False
            else:
                days_per_label = 1 # set to 1 and flex_on to True to have some flexibilty in x_labels
            while m < len(the_days):
                day_labels.append('%s %s' % (str(d), mth_labels[m]))
                d += days_per_label
                if d > the_days[m]:
                    d = d - the_days[m]
                    m += 1
            x = []
            len_x = self.rows - self.toprow[1]
            if self.plottype.currentText() == 'Step Day Chart':
                len_x = int((len_x + self.interval - 1) / self.interval)
            for i in range(len_x):
                x.append(i)
            load = []
            tgt_col = -1
            overlay_cols = []
            overlay = []
            for o in range(self.no_of_overlays):
                overlay_cols.append([])
                overlay.append([])
            data = []
            maxy = 0
            miny = 0
            titl = self.replace_words('y', self.title.text(), str(year))
            titl = self.replace_words('m', titl, '')
            titl = titl.replace('  ', ' ')
            titl = titl.replace('Diurnal ', '')
            titl = titl.replace('Diurnal', '')
            titl = self.replace_words('s', titl, isheet)
            if suptitle != '':
                suptitle = self.replace_words('y', suptitle, str(year))
                suptitle = self.replace_words('m', suptitle, '')
                suptitle = suptitle.replace('  ', ' ')
                suptitle = suptitle.replace('Diurnal ', '')
                suptitle = suptitle.replace('Diurnal', '')
                suptitle = self.replace_words('s', suptitle, isheet)
            for c2 in range(2, ws.ncols):
                try:
                    column = ws.cell_value(self.toprow[0], c2).replace('\n',' ')
                except:
                    column = str(ws.cell_value(self.toprow[0], c2))
                if self.zone_row > 0 and ws.cell_value(self.zone_row, c2) != '' and ws.cell_value(self.zone_row, c2) is not None:
                    column = ws.cell_value(self.zone_row, c2).replace('\n',' ') + '.' + column
                if column == self.target:
                    tgt_col = c2
                for o in range(self.no_of_overlays):
                    if self.overlay[o] != '<none>':
                        if column[:len(self.overlay[o])] == self.overlay[o] or \
                          ((self.overlay[o] == 'Charge' or self.overlay[o] == 'Underlying Load') and column == self.target):
                            overlay_cols[o].append(c2)
            if self.plottype.currentText() != 'Step Day Chart':
                if len(breakdowns) > 0:
                    for c in range(self.order.count() -1, -1, -1):
                        col = self.order.item(c).text()
                        for c2 in range(2, ws.ncols):
                            try:
                                brkdown = ws.cell_value(self.breakdown_row, c2)
                                if brkdown is not None and brkdown != '' and brkdown != breakdowns[0]:
                                    continue
                            except:
                                pass
                            try:
                                column = ws.cell_value(self.toprow[0], c2).replace('\n',' ')
                            except:
                                column = str(ws.cell_value(self.toprow[0], c2))
                            if self.zone_row > 0 and ws.cell_value(self.zone_row, c2) != '' and ws.cell_value(self.zone_row, c2) is not None:
                                column = ws.cell_value(self.zone_row, c2).replace('\n',' ') + '.' + column
                            if column == col:
                                data.append([])
                                label.append(column)
                                for row in range(self.toprow[1] + 1, self.rows + 1):
                                    data[-1].append(ws.cell_value(row, c2))
                                    try:
                                        maxy = max(maxy, data[-1][-1])
                                        miny = min(miny, data[-1][-1])
                                    except:
                                        self.log.setText(f'Data error with {column} ({ssCol(c2, base=0)}{row + 1}). Period may be incomplete (1)')
                                        return
                    for breakdown in breakdowns[1:]:
                        for c in range(self.order.count() -1, -1, -1):
                            col = self.order.item(c).text()
                            for c2 in range(2, ws.ncols):
                                try:
                                    brkdown = ws.cell_value(self.breakdown_row, c2)
                                    if brkdown != breakdown:
                                        continue
                                except:
                                    continue
                                try:
                                    column = ws.cell_value(self.toprow[0], c2).replace('\n',' ')
                                except:
                                    column = str(ws.cell_value(self.toprow[0], c2))
                                if self.zone_row > 0 and ws.cell_value(self.zone_row, c2) != '' and ws.cell_value(self.zone_row, c2) is not None:
                                    column = ws.cell_value(self.zone_row, c2).replace('\n',' ') + '.' + column
                                if column == col:
                                    data.append([])
                                    label.append(self.short_legend + column + ' ' + breakdown)
                                    for row in range(self.toprow[1] + 1, self.rows + 1):
                                        data[-1].append(ws.cell_value(row, c2))
                                        try:
                                            maxy = max(maxy, data[-1][-1])
                                            miny = min(miny, data[-1][-1])
                                        except:
                                            self.log.setText(f'Data error with {column} ({ssCol(c2, base=0)}{row + 1}). Period may be incomplete (2)')
                                            return
                else:
                    for c in range(self.order.count() -1, -1, -1):
                        col = self.order.item(c).text()
                        for c2 in range(2, ws.ncols):
                            try:
                                column = ws.cell_value(self.toprow[0], c2).replace('\n',' ')
                            except:
                                column = str(ws.cell_value(self.toprow[0], c2))
                            if self.zone_row > 0 and ws.cell_value(self.zone_row, c2) != '' and ws.cell_value(self.zone_row, c2) is not None:
                                column = ws.cell_value(self.zone_row, c2).replace('\n',' ') + '.' + column
                            if column == col:
                                data.append([])
                                label.append(column)
                                for row in range(self.toprow[1] + 1, self.rows + 1):
                                    data[-1].append(ws.cell_value(row, c2))
                                    try:
                                        maxy = max(maxy, data[-1][-1])
                                        miny = min(miny, data[-1][-1])
                                    except:
                                        self.log.setText(f'Data error with {column} ({ssCol(c2, base=0)}{row + 1}). Period may be incomplete (3)')
                                        return
                                break
                if tgt_col >= 0:
                    for row in range(self.toprow[1] + 1, self.rows + 1):
                        load.append(ws.cell_value(row, tgt_col))
                        maxy = max(maxy, load[-1])
                        miny = min(miny, load[-1])
                for o in range(self.no_of_overlays):
                    if len(overlay_cols[o]) > 0:
                        for row in range(self.toprow[1] + 1, self.rows + 1):
                            overlay[o].append(0.)
                        if self.overlay[o] == 'Charge':
                            for row in range(self.toprow[1] + 1, self.rows + 1):
                                for col in overlay_cols[o]:
                                    overlay[o][row - self.toprow[1] - 1] += ws.cell_value(row, col)
                        elif self.overlay[o] == 'Underlying Load':
                            for row in range(self.toprow[1] + 1, self.rows + 1):
                                overlay[o][row - self.toprow[1] - 1] = ws.cell_value(row, overlay_cols[o][-1])
                        else:
                            for row in range(self.toprow[1] + 1, self.rows + 1):
                                overlay[o][row - self.toprow[1] - 1] = ws.cell_value(row, overlay_cols[o][-1])
                        for h in range(len(overlay[o])):
                            maxy = max(maxy, overlay[o][h])
                            miny = min(miny, overlay[o][h])
            else: # 'Step Day Chart'
                if len(breakdowns) > 0:
                    for c in range(self.order.count() -1, -1, -1):
                        col = self.order.item(c).text()
                        for c2 in range(2, ws.ncols):
                            try:
                                brkdown = ws.cell_value(self.breakdown_row, c2)
                                if brkdown is not None and brkdown != '' and brkdown != breakdowns[0]:
                                    continue
                            except:
                                pass
                            try:
                                column = ws.cell_value(self.toprow[0], c2).replace('\n',' ')
                            except:
                                column = str(ws.cell_value(self.toprow[0], c2))
                            if self.zone_row > 0 and ws.cell_value(self.zone_row, c2) != '' and ws.cell_value(self.zone_row, c2) is not None:
                                column = ws.cell_value(self.zone_row, c2).replace('\n',' ') + '.' + column
                            if column == col:
                                data.append([])
                                label.append(column)
                                for row in range(self.toprow[1] + 1, self.rows + 1, self.interval):
                                    data[-1].append(0)
                                    for r2 in range(row, row + self.interval):
                                        data[-1][-1] += ws.cell_value(r2, c2)
                                    try:
                                        maxy = max(maxy, data[-1][-1])
                                        miny = min(miny, data[-1][-1])
                                    except:
                                        self.log.setText(f'Data error with {column} ({ssCol(c2, base=0)}{row + 1}). Period may be incomplete (4)')
                                        return
                    for breakdown in breakdowns[1:]:
                        for c in range(self.order.count() -1, -1, -1):
                            col = self.order.item(c).text()
                            for c2 in range(2, ws.ncols):
                                try:
                                    brkdown = ws.cell_value(self.breakdown_row, c2)
                                    if brkdown != breakdown:
                                        continue
                                except:
                                    continue
                                try:
                                    column = ws.cell_value(self.toprow[0], c2).replace('\n',' ')
                                except:
                                    column = str(ws.cell_value(self.toprow[0], c2))
                                if self.zone_row > 0 and ws.cell_value(self.zone_row, c2) != '' and ws.cell_value(self.zone_row, c2) is not None:
                                    column = ws.cell_value(self.zone_row, c2).replace('\n',' ') + '.' + column
                                if column == col:
                                    data.append([])
                                    label.append(self.short_legend + column + ' ' + breakdown)
                                    for row in range(self.toprow[1] + 1, self.rows + 1, self.interval):
                                        data[-1].append(0)
                                        for r2 in range(row, row + self.interval):
                                            data[-1][-1] += ws.cell_value(r2, c2)
                                        try:
                                            maxy = max(maxy, data[-1][-1])
                                            miny = min(miny, data[-1][-1])
                                        except:
                                            self.log.setText(f'Data error with {column} ({ssCol(c2, base=0)}{row + 1}). Period may be incomplete (5)')
                                            return
                else:
                    for c in range(self.order.count() -1, -1, -1):
                        col = self.order.item(c).text()
                        for c2 in range(2, ws.ncols):
                            try:
                                column = ws.cell_value(self.toprow[0], c2).replace('\n',' ')
                            except:
                                column = str(ws.cell_value(self.toprow[0], c2))
                            if self.zone_row > 0 and ws.cell_value(self.zone_row, c2) != '' and ws.cell_value(self.zone_row, c2) is not None:
                                column = ws.cell_value(self.zone_row, c2).replace('\n',' ') + '.' + column
                            if column == col:
                                data.append([])
                                label.append(column)
                                for row in range(self.toprow[1] + 1, self.rows + 1, self.interval):
                                    data[-1].append(0)
                                    for r2 in range(row, row + self.interval):
                                        try:
                                            data[-1][-1] += ws.cell_value(r2, c2)
                                        except:
                                            if r2 >= ws.nrows:
                                                break
                                            self.log.setText(f'Data error with {column} ({ssCol(c2, base=0)}{r2 + 1}). Period may be incomplete (6)')
                                            return
                                    maxy = max(maxy, data[-1][-1])
                                    miny = min(miny, data[-1][-1])
                                break
                if tgt_col >= 0:
                    for row in range(self.toprow[1] + 1, self.rows + 1, self.interval):
                        load.append(0)
                        for r2 in range(row, row + self.interval):
                            load[-1] += ws.cell_value(r2, tgt_col)
                        maxy = max(maxy, load[-1])
                        miny = min(miny, load[-1])
                for o in range(self.no_of_overlays):
                    if len(overlay_cols[o]) > 0:
                        if self.overlay[o] == 'Charge':
                            for row in range(self.toprow[1] + 1, self.rows + 1, self.interval):
                                overlay[o].append(0.)
                                for col in overlay_cols[o]:
                                    for r2 in range(row, row + self.interval):
                                        overlay[o][-1] += ws.cell_value(r2, col)
                        elif self.overlay[o] == 'Underlying Load':
                            for row in range(self.toprow[1] + 1, self.rows + 1, self.interval):
                                overlay[o].append(0.)
                                for r2 in range(row, row + self.interval):
                                    overlay[o][-1] += ws.cell_value(row, overlay_cols[o][-1])
                        else:
                            if self.overlay_type[o] in ['Average', 'Sum']:
                                for row in range(self.toprow[1] + 1, self.rows + 1, self.interval):
                                    overlay[o].append(0.)
                                    for r2 in range(row, row + self.interval):
                                        try:
                                            overlay[o][-1] += ws.cell_value(r2, overlay_cols[o][-1])
                                        except:
                                            self.log.setText(f'Data error with {column} ({ssCol(c2, base=0)}{row + 1}). Period may be incomplete (7)')
                                            return
                                    if self.overlay_type[o] == 'Average':
                                        overlay[o][-1] = overlay[o][-1] / self.interval
                            else:
                                h = int(self.overlay_type[:2])
                                for row in range(self.toprow[1] + 1, self.rows + 1, self.interval):
                                    overlay[o].append(ws.cell_value(row + h, overlay_cols[o][-1]))
                        for h in range(len(overlay[o])):
                            maxy = max(maxy, overlay[o][h])
                            miny = min(miny, overlay[o][h])
            if do_12:
                ## to get 12 months on a page
                matplotlib.rcParams['figure.figsize'] = [18, 2.2]
            if self.plottype.currentText() == 'Line Chart' or self.plottype.currentText() == 'Step Line Chart' \
              or self.plottype.currentText() == 'Step Day Chart':
                fig = plt.figure(figname, constrained_layout=self.constrained_layout)
                if suptitle != '':
                    fig.suptitle(suptitle, fontproperties=self.fontprops['Header'])
                if gridtype != '':
                    plt.grid(axis=gridtype)
                lc1 = plt.subplot(111)
                if self.legend_side == 'Right':
                    plt.subplots_adjust(left=0.1, bottom=0.1, right=0.75)
                elif self.legend_side == 'Left':
                    plt.subplots_adjust(left=0.25, bottom=0.1, right=0.9)
                    loc = 'lower left'
                plt.title(titl, fontdict=self.fontprops['Title'])
                if self.plottype.currentText() == 'Line Chart':
                    for c in range(len(data)):
                        lc1.plot(x, data[c], linewidth=self.linSpin.value(), label=label[c], color=self.set_colour(label[c]))
                    if len(load) > 0:
                        lc1.plot(x, load, linewidth=self.tgtSpin.value(), label=self.short_legend + self.target,
                                 color=self.set_colour(self.target), linestyle=self.tgtLine.currentText())
                    for o in range(self.no_of_overlays):
                        if len(overlay[o]) > 0:
                            lc1.plot(x, overlay[o], linewidth=self.ovrSpin[o].value(), label=self.short_legend + self.overlay[o],
                                     color=self.overlay_colour[o], linestyle=self.ovrLine[o].currentText())
                else:
                    for c in range(len(data)):
                        lc1.step(x, data[c], linewidth=self.linSpin.value(), label=label[c],
                                 color=self.set_colour(label[c]), linestyle=self.linLine.currentText())
                    if len(load) > 0:
                        lc1.step(x, load, linewidth=self.tgtSpin.value(), label=self.short_legend + self.target,
                                 color=self.set_colour(self.target), linestyle=self.tgtLine.currentText())
                    for o in range(self.no_of_overlays):
                        if len(overlay[o]) > 0:
                            lc1.step(x, overlay[o], linewidth=self.ovrSpin[o].value(), label=self.short_legend + self.overlay[o],
                                     color=self.overlay_colour[o], linestyle=self.ovrLine[o].currentText())
                yminmax = set_ylimits(miny, maxy)
                lc1.legend(bbox_to_anchor=[0.5, -0.1], loc='center', ncol=(len(data) + 1),
                           prop=self.legend_font)
                plt.ylim(yminmax)
                plt.xlim([0, len(x)])
                if not do_12:
                    if self.plottype.currentText() != 'Step Day Chart':
                        xticks = list(range(0, len(x), self.interval * days_per_label))
                        xticklabels = day_labels[:len(xticks)]
                    else:
                        xticks = [0]
                        xticklabels = [mth_labels[0]]
                        for m in range(len(the_days) - 1):
                            xticks.append(xticks[-1] + the_days[m])
                            xticklabels.append(mth_labels[m + 1])
                    lc1.set_xticks(xticks)
                    if len(xticks) == len(xticklabels) + 1:
                        xticklabels.append('?')
                    lc1.set_xticklabels(xticklabels, rotation='vertical', fontdict=self.fontprops['Ticks'])
                    lc1.set_xlabel('Period', fontdict=self.fontprops['Label'])
                lc1.tick_params(colors=self.fontprops['Ticks']['color'], which='both')
                lc1.set_ylabel('Power (MW/MWh)', fontdict=self.fontprops['Label']) # MWh?
                if self.y2_label != '':
                    secax = lc1.secondary_yaxis(location='right', functions=(self.transform_y2, self.inverse_y2))
                    secax.set_ylabel(self.y2_label)
                zp = ZoomPanX()
                f = zp.zoom_pan(lc1, base_scale=1.2, flex_ticks=flex_on, mth_labels=do_12_labels) # enable scrollable zoom
                plt.show()
                del zp
            elif self.plottype.currentText() in ['Cumulative', 'Step Chart']:
                if self.plottype.currentText() == 'Cumulative':
                    step = None
                else:
                    step = 'pre'
                fig = plt.figure(figname, constrained_layout=self.constrained_layout)
                if suptitle != '':
                    fig.suptitle(suptitle, fontproperties=self.fontprops['Header'])
                if gridtype != '':
                    plt.grid(axis=gridtype)
                cu1 = plt.subplot(111)
                if self.legend_side == 'Right':
                    plt.subplots_adjust(left=0.1, bottom=0.1, right=0.75)
                elif self.legend_side == 'Left':
                    plt.subplots_adjust(left=0.25, bottom=0.1, right=0.9)
                    loc = 'lower left'
                if not do_12:
                    plt.title(titl, fontdict=self.fontprops['Title'])
                if self.percentage.isChecked():
                    totals = [0.] * len(x)
                    bottoms = [0.] * len(x)
                    values = [0.] * len(x)
                    for c in range(len(data)):
                        for h in range(len(data[c])):
                            totals[h] = totals[h] + data[c][h]
                    for h in range(len(data[0])):
                        values[h] = data[0][h] / totals[h] * 100.
                    cu1.fill_between(x, 0, values, label=label[0], color=self.set_colour(label[0]), hatch=self.set_hatch(label[0]),
                                     alpha=self.alpha_fill, step=step)
                    for c in range(1, len(data)):
                        for h in range(len(data[c])):
                            bottoms[h] = values[h]
                            values[h] = values[h] + data[c][h] / totals[h] * 100.
                        cu1.fill_between(x, bottoms, values, label=label[c], color=self.set_colour(label[c]), hatch=self.set_hatch(label[c]),
                                         alpha=self.alpha_fill, step=step)
                    maxy = 100
                    cu1.set_ylabel('Power (%)', fontdict=self.fontprops['Label'])
                else:
                    if self.target == '<none>':
                        cu1.fill_between(x, 0, data[0], label=label[0], color=self.set_colour(label[0]), hatch=self.set_hatch(label[0]),
                                         alpha=self.alpha_fill, step=step)
                        for c in range(1, len(data)):
                            for h in range(len(data[c])):
                                data[c][h] = data[c][h] + data[c - 1][h]
                                maxy = max(maxy, data[c][h])
                            cu1.fill_between(x, data[c - 1], data[c], label=label[c], color=self.set_colour(label[c]), hatch=self.set_hatch(label[c]),
                                             alpha=self.alpha_fill, step=step)
                        top = data[0][:]
                        for d in range(1, len(data)):
                            for h in range(len(top)):
                                top[h] = max(top[h], data[d][h])
                        if self.plottype.currentText() == 'Cumulative':
                            cu1.plot(x, top, color='white')
                        else:
                            cu1.step(x, top, color='white')
                    else:
                 #       pattern = ['-', '+', 'x', '\\', '|', '/', '*', 'o', 'O', '.']
                 #       pat = 0
                        full = []
                        for h in range(len(load)):
                           full.append(min(load[h], data[0][h]))
                        cu1.fill_between(x, 0, full, label=label[0], color=self.set_colour(label[0]), hatch=self.set_hatch(label[0]),
                                         alpha=self.alpha_fill, step=step)
                        for h in range(len(data[0])):
                            if data[0][h] > full[h]:
                                if self.spill_label.text() != '':
                                    cu1.fill_between(x, full, data[0], alpha=self.alpha, color=self.set_colour(label[0]), hatch=self.set_hatch(label[0]),
                                        label=label[0] + ' ' + self.spill_label.text(), step=step)
                                else:
                                    cu1.fill_between(x, full, data[c], alpha=self.alpha, color=self.set_colour(label[c]), hatch=self.set_hatch(label[c]),
                                                     step=step)
                                break
                        for c in range(1, len(data)):
                            full = []
                            for h in range(len(data[c])):
                                data[c][h] = data[c][h] + data[c - 1][h]
                                maxy = max(maxy, data[c][h])
                                full.append(max(min(load[h], data[c][h]), data[c - 1][h]))
                            cu1.fill_between(x, data[c - 1], full, label=label[c], color=self.set_colour(label[c]), hatch=self.set_hatch(label[c]),
                                             alpha=self.alpha_fill, step=step)
                            for h in range(len(data[c])):
                                if data[c][h] > full[h] + self.margin_of_error:
                                    if self.spill_label.text() != '':
                                        cu1.fill_between(x, full, data[c], color=self.set_colour(label[c]), hatch=self.set_hatch(label[c]),
                                                         alpha=self.alpha_fill, label=label[c] + ' ' + self.spill_label.text(), step=step)
                                    else:
                                        cu1.fill_between(x, full, data[c], alpha=self.alpha, color=self.set_colour(label[c]), hatch=self.set_hatch(label[c]),
                                                         step=step)
                                    break
                        top = data[0][:]
                        for d in range(1, len(data)):
                            for h in range(len(top)):
                                top[h] = max(top[h], data[d][h])
                        if self.plottype.currentText() == 'Cumulative':
                            if self.alpha == 0:
                                cu1.plot(x, top, color='gray', linestyle='dashed')
                            else:
                                cu1.plot(x, top, color='gray')
                        else:
                            if self.alpha == 0:
                                cu1.step(x, top, color='gray', linestyle='dashed')
                            else:
                                cu1.step(x, top, color='gray')
                        short = []
                        do_short = False
                        c = len(data) - 1
                        for h in range(len(load)):
                            if load[h] > data[c][h] + self.margin_of_error:
                                do_short = True
                            short.append(max(data[c][h], load[h]))
                        if do_short:
                            if self.colours['shortfall'][0].lower() != 'w' and self.colours['shortfall'].lower() != '#ffffff':
                                cu1.fill_between(x, data[c], short, label='Shortfall', color=self.set_colour('shortfall'),
                                                 alpha=self.alpha_fill, step=step)
                        if self.plottype.currentText() == 'Cumulative':
                            cu1.plot(x, load, linewidth=self.tgtSpin.value(), label=self.short_legend + self.target,
                                     color=self.set_colour(self.target), linestyle=self.tgtLine.currentText())
                        else:
                            cu1.step(x, load, linewidth=self.tgtSpin.value(), label=self.short_legend + self.target,
                                     color=self.set_colour(self.target), linestyle=self.tgtLine.currentText())
                    for o in range(self.no_of_overlays):
                        if len(overlay[o]) > 0:
                            if self.plottype.currentText() == 'Cumulative':
                                cu1.plot(x, overlay[o], linewidth=self.ovrSpin[o].value(), label=self.short_legend + self.overlay[o],
                                         color=self.overlay_colour[o], linestyle=self.ovrLine[o].currentText())
                            else:
                                cu1.step(x, overlay[o], linewidth=self.ovrSpin[o].value(), label=self.short_legend + self.overlay[o],
                                         color=self.overlay_colour[o], linestyle=self.ovrLine[o].currentText())
                    cu1.set_ylabel('Power (MW/MWh)', fontdict=self.fontprops['Label'])
        #        miny = 0
                yminmax = set_ylimits(miny, maxy)
                cu1.legend(bbox_to_anchor=[0.5, -0.1], loc='center', ncol=(len(data) + 2), prop=self.legend_font)
                plt.ylim(yminmax)
                plt.xlim([0, len(x)])
                if not do_12:
                    xticks = list(range(0, len(x), self.interval * days_per_label))
                    cu1.set_xticks(xticks)
                    cu1.set_xticklabels(day_labels[:len(xticks)], rotation='vertical', fontdict=self.fontprops['Ticks'])
                    cu1.set_xlabel('Period', fontdict=self.fontprops['Label'])
                cu1.tick_params(colors=self.fontprops['Ticks']['color'], which='both')
                if self.y2_label != '':
                    secax = cu1.secondary_yaxis(location='right', functions=(self.transform_y2, self.inverse_y2))
                    secax.set_ylabel(self.y2_label)
                zp = ZoomPanX()
                f = zp.zoom_pan(cu1, base_scale=1.2, flex_ticks=flex_on, mth_labels=do_12_labels) # enable scrollable zoom
                plt.show()
                del zp
            elif self.plottype.currentText() == 'Bar Chart':
                msgbox = QtWidgets.QMessageBox()
                msgbox.setWindowTitle('powerplot - Bar Chart')
                msgbox.setText("This could take a while to display. Are you sure?")
                msgbox.setIcon(QtWidgets.QMessageBox.Question)
                msgbox.setStandardButtons(QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
                reply = msgbox.exec_()
                if reply != QtWidgets.QMessageBox.Yes:
                    return
                fig = plt.figure(figname, constrained_layout=self.constrained_layout)
                if suptitle != '':
                    fig.suptitle(suptitle, fontproperties=self.fontprops['Header'])
                if gridtype != '':
                    plt.grid(axis=gridtype)
                bc1 = plt.subplot(111)
                if self.legend_side == 'Right':
                    plt.subplots_adjust(left=0.1, bottom=0.1, right=0.75)
                elif self.legend_side == 'Left':
                    plt.subplots_adjust(left=0.25, bottom=0.1, right=0.9)
                    loc = 'lower left'
                if not do_12:
                    plt.title(titl, fontdict=self.fontprops['Title'])
                if self.percentage.isChecked():
                    miny = 0
                    totals = [0.] * len(x)
                    bottoms = [0.] * len(x)
                    values = [0.] * len(x)
                    for c in range(len(data)):
                        for h in range(len(data[c])):
                            totals[h] = totals[h] + data[c][h]
                    for h in range(len(data[0])):
                        values[h] = data[0][h] / totals[h] * 100.
                    bc1.bar(x, values, label=label[0], color=self.set_colour(label[0]), alpha=self.alpha_fill,
                            hatch=self.set_hatch(label[0]))
                    for c in range(1, len(data)):
                        for h in range(len(data[c])):
                            bottoms[h] = bottoms[h] + values[h]
                            values[h] = data[c][h] / totals[h] * 100.
                        bc1.bar(x, values, bottom=bottoms, label=label[c], color=self.set_colour(label[c]), alpha=self.alpha_fill,
                                hatch=self.set_hatch(label[c]))
                    maxy = 100
                    bc1.set_ylabel('Power (%)', fontdict=self.fontprops['Label'])
                else:
                    if self.target == '<none>':
                        bc1.bar(x, data[0], label=label[0], color=self.set_colour(label[0]), alpha=self.alpha_fill,
                                hatch=self.set_hatch(label[0]))
                        bottoms = [0.] * len(x)
                        for c in range(1, len(data)):
                            for h in range(len(data[c])):
                                bottoms[h] = bottoms[h] + data[c - 1][h]
                                maxy = max(maxy, data[c][h])
                            bc1.bar(x, data[c], bottom=bottoms, label=label[c], color=self.set_colour(label[c]), alpha=self.alpha_fill,
                                    hatch=self.set_hatch(label[c]))
                    else:
                        bc1.bar(x, data[0], label=label[0], color=self.set_colour(label[0]), alpha=self.alpha_fill,
                                hatch=self.set_hatch(label[0]))
                        bottoms = [0.] * len(x)
                        for c in range(1, len(data)):
                            for h in range(len(data[c])):
                                bottoms[h] = bottoms[h] + data[c - 1][h]
                                maxy = max(maxy, data[c][h])
                            bc1.bar(x, data[c], bottom=bottoms, label=label[c], color=self.set_colour(label[c]), alpha=self.alpha_fill,
                                    hatch=self.set_hatch(label[c]))
                        bc1.plot(x, load, linewidth=self.tgtSpin.value(), label=self.short_legend + self.target, color=self.set_colour(self.target),
                                 linestyle=self.tgtLine.currentText())
     # bug here i think
                        for o in range(self.no_of_overlays):
                            bc1.plot(x, overlay[o], linewidth=self.ovrSpin[o].value(), label=self.short_legend + self.overlay[o],
                                     color=self.overlay_colour[o], linestyle=self.ovrLine[o].currentText())
                    bc1.set_ylabel('Power (MW/MWh)', fontdict=self.fontprops['Label'])
        #        miny = 0
                yminmax = set_ylimits(miny, maxy)
                bc1.legend(bbox_to_anchor=[0.5, -0.1], loc='center', ncol=(len(data) + 2), prop=self.legend_font)
                plt.ylim(yminmax)
                plt.xlim([0, len(x)])
                if not do_12:
                    xticks = list(range(0, len(x), self.interval * days_per_label))
                    bc1.set_xticks(xticks)
                    bc1.set_xticklabels(day_labels[:len(xticks)], rotation='vertical', fontdict=self.fontprops['Ticks'])
                    bc1.set_xlabel('Period', fontdict=self.fontprops['Label'])
                bc1.tick_params(colors=self.fontprops['Ticks']['color'], which='both')
                if self.y2_label != '':
                    secax = bc1.secondary_yaxis(location='right', functions=(self.transform_y2, self.inverse_y2))
                    secax.set_ylabel(self.y2_label)
                zp = ZoomPanX()
                f = zp.zoom_pan(bc1, base_scale=1.2, flex_ticks=flex_on, mth_labels=do_12_labels) # enable scrollable zoom
                plt.show()
                del zp
            elif self.plottype.currentText() == 'Heat Map':
                fig = plt.figure(figname)
                if suptitle != '':
                    fig.suptitle(suptitle, fontproperties=self.fontprops['Header'])
                if gridtype != '':
                    plt.grid(axis=gridtype)
                hm1 = plt.subplot(111)
                if not do_12:
                    plt.title(titl, fontdict=self.fontprops['Title'])
                hmdata = []
                for hr in range(self.interval):
                    hmdata.append([])
                    for dy in range(365 + int(self.leapyear)):
                        hmdata[-1].append(0)
                for col in range(len(data)):
                    x = 0
                    y = -1
                    for hr in range(len(data[col])):
                        y += 1
                        if y > self.interval - 1:
                            x += 1
                            y = 0
                        hmdata[y][x] += data[col][hr]
                if tgt_col >= 0:
                    x = 0
                    y = -1
                    for hr in range(len(load)):
                        y += 1
                        if y > self.interval - 1:
                            x += 1
                            y = 0
                        try:
                            hmdata[y][x] = hmdata[y][x] / load[hr]
                        except:
                            pass
                miny = 999999
                maxy = 0
                for y in range(len(hmdata)):
                    for x in range(len(hmdata[0])):
                    #    hmdata[y][x] = round(hmdata[y][x], 3)
                        miny = min(miny, hmdata[y][x])
                        maxy = max(maxy, hmdata[y][x])
                # print(hmdata)
                if tgt_col >= 0:
                    fmt_str = '{:.3f}'
                    vmax = 1
                    vmin = 0
                    if maxy > 1:
                        maxy = 1
                else:
                    fmt_str = '{:,.0f}'
                    if self.cbar2:
                        vmax = maxy
                        vmin = miny
                    else:
                        vmax = None
                        vmin = None
                im1 = hm1.imshow(hmdata, cmap=self.cmap, interpolation='nearest', aspect='auto', vmax=vmax, vmin=vmin)
                day = -0.5
                xticks = [day]
                for m in range(len(the_days) - 1):
                    day += the_days[m]
                    xticks.append(day)
                hm1.set_xticks(xticks)
                hm1.set_xticklabels(mth_labels, ha='left', fontdict=self.fontprops['Ticks'])
                yticks = [-0.5]
                if self.interval == 24:
                    for y in range(0, 24, 4):
                        yticks.append(y + 2.5)
                else:
                    for y in range(0, 48, 8):
                        yticks.append(y + 6.5)
                hm1.invert_yaxis()
                hm1.set_yticks(yticks)
                hm1.set_yticklabels(hr_labels, va='bottom', fontdict=self.fontprops['Ticks'])
                hm1.set_xlabel('Period', fontdict=self.fontprops['Label'])
                hm1.set_ylabel('Hour', fontdict=self.fontprops['Label'])
                if self.cbar:
                    fig.subplots_adjust(bottom=0.1, right=0.8, top=0.9)
                    cb_ax = fig.add_axes([0.85, 0.1, 0.025, 0.8])
                    cbar = fig.colorbar(im1, cax=cb_ax)
                    if self.cbar2:
                        curcticks = cbar.get_ticks()
                        cticks = [vmin, vmax]
                        if miny != vmin:
                            cticks.append(miny)
                        if maxy != vmax:
                            cticks.append(maxy)
                        for y in range(len(curcticks)):
                            if curcticks[y] > vmin and curcticks[y] < vmax:
                                cticks.append(curcticks[y])
                        cticks.sort()
                        try:
                            cbar.set_ticks(cticks, fontdict=self.fontprops['Ticks'])
                        except:
                            clabels = []
                            for ct in cticks:
                                clabels.append(f'{ct}')
                            cbar.set_ticks(cticks, labels=clabels, fontdict=self.fontprops['Ticks'])
                self.log.setText('Heat map value range: {} to {}'.format(fmt_str, fmt_str).format(miny, maxy))
                QtCore.QCoreApplication.processEvents()
                plt.show()
            elif self.plottype.currentText() == 'Load Duration':
                fig = plt.figure(figname, constrained_layout=self.constrained_layout)
                if suptitle != '':
                    fig.suptitle(suptitle, fontproperties=self.fontprops['Header'])
                plt.title(titl, fontdict=self.fontprops['Title'])
                if gridtype != '':
                    plt.grid(axis=gridtype)
                ld1 = plt.subplot(111)
                if self.legend_side == 'Right':
                    plt.subplots_adjust(left=0.1, bottom=0.1, right=0.75)
                elif self.legend_side == 'Left':
                    plt.subplots_adjust(left=0.25, bottom=0.1, right=0.9)
                    loc = 'lower left'
                miny = 0
                maxy = 0
                for c in range(len(data)):
                    data[c] = sorted(data[c], reverse=True)
                    maxy = max(maxy, data[c][0])
                    ld1.plot(x, data[c], label=label[c], color=self.set_colour(label[c]))
                if tgt_col >= 0:
                    if len(load) == 0:
                        for row in range(self.toprow[1] + 1, self.rows + 1):
                            load.append(ws.cell_value(row, tgt_col))
                    load = sorted(load, reverse=True)
                    maxy = max(maxy, load[0])
                    ld1.plot(x, load, linewidth=self.tgtSpin.value(), label=self.short_legend + self.target,
                             color=self.set_colour(self.target), linestyle=self.tgtLine.currentText())
                yminmax = set_ylimits(miny, maxy)
                ld1.legend(bbox_to_anchor=[0.5, -0.1], loc='center', ncol=(len(data) + 2), prop=self.legend_font)
                plt.ylim(yminmax)
                plt.xlim([0, len(x)])
                xticks = []
                xticklabels = []
                if len(x) > 1000:
                    incr = 1000
                else:
                    incr = 100
                for i in range(0, len(x), incr):
                    xticks.append(i)
                    xticklabels.append(f'{i}')
                ld1.set_xticks(xticks)
                ld1.set_xticklabels(xticklabels, fontdict=self.fontprops['Ticks'])
                ld1.set_xlabel('Power in descending order', fontdict=self.fontprops['Label'])
                ld1.set_ylabel('Power (MW/MWh)', fontdict=self.fontprops['Label']) # MWh?
                ld1.tick_params(colors=self.fontprops['Ticks']['color'], which='both')
                zp = ZoomPanX()
                f = zp.zoom_pan(ld1, base_scale=1.2) #, flex_ticks=flex_on)
                plt.show()
                del zp
            if do_12:
                matplotlib.rcParams['figure.figsize'] = do_12_save
        else: # diurnal average - period chosen
            # need to check as 1 day is wrong if target set
            if self.interval == 24:
                ticks = list(range(0, 21, 4))
                ticks.append(23)
            else:
                ticks = list(range(0, 41, 8))
                ticks.append(47)
            titl = self.title.text()
            titl = self.replace_words('s', self.title.text(), isheet)
            if suptitle != '':
                suptitle = self.replace_words('s', suptitle, isheet)
            if self.select_day and self.aday.currentText() != '':
                if self.period.currentText() == 'Year':
                    i = 0
                else:
                    try:
                        i = self.seasons[self.period.currentText()][0]
                    except:
                        i = mth_labels.index(self.period.currentText())
                if self.aday.currentIndex() > the_days[i]:
                    self.aday.setCurrentIndex(self.aday.currentIndex() - the_days[i])
                    i += 1
                self.period.setCurrentIndex(i + 2)
                d = 0
                for m in range(i):
                    d += the_days[m]
                d += self.aday.currentIndex() - 1
                strt_row = [d * self.interval + self.toprow[1]]
                todo_rows = [self.interval]
                y = self.replace_words('y', self.title.text(), 'find')
                m = self.replace_words('m', self.title.text(), 'find')
                if y > m:
                    txt = self.aday.currentText() + ' ' + self.period.currentText()
                else:
                    txt = self.period.currentText() + ' ' + self.aday.currentText()
                titl = self.replace_words('y', self.title.text(), str(year))
                titl = self.replace_words('m', titl, txt)
                if suptitle != '':
                    suptitle = self.replace_words('y', suptitle, str(year))
                    suptitle = self.replace_words('m', suptitle, txt)
            else:
                titl = self.replace_words('y', self.title.text(), str(year))
                if suptitle != '':
                    suptitle = self.replace_words('y', suptitle, str(year))
                todo_mths = []
                if self.period.currentText() == 'Year':
                    titl = self.replace_words('m', titl, '')
                    titl = titl.replace('  ', ' ')
                    if suptitle != '':
                        suptitle = self.replace_words('m', suptitle, '')
                        suptitle = suptitle.replace('  ', ' ')
                    strt_row = [self.toprow[1]]
                    todo_rows = [self.rows - strt_row[0]]
                else:
                    strt_row = []
                    todo_rows = []
                    if self.period.currentText() in self.seasons.keys():
                        titl = self.replace_words('m', titl, self.period.currentText())
                        if suptitle != '':
                            suptitle = self.replace_words('m', suptitle, '')
                        for s in self.seasons[self.period.currentText()]:
                            m = 0
                            strt_row.append(0)
                            while m < s:
                                strt_row[-1] = strt_row[-1] + the_days[m] * self.interval
                                m += 1
                            todo_mths.append(m)
                            strt_row[-1] = strt_row[-1] + self.toprow[1]
                            todo_rows.append(the_days[s] * self.interval)
                    else:
                         i = mth_labels.index(self.period.currentText())
                         todo_mths = [i]
                         if self.cperiod.currentText() == '<none>':
                             titl = self.replace_words('m', titl, self.period.currentText())
                             if suptitle != '':
                                 suptitle = self.replace_words('m', suptitle, self.period.currentText())
                         else:
                             titl = self.replace_words('m', titl, self.period.currentText() + \
                                    ' to ' + self.cperiod.currentText())
                             if suptitle != '':
                                 suptitle = self.replace_words('m', suptitle, self.period.currentText() + \
                                            ' to ' + self.cperiod.currentText())
                             j = mth_labels.index(self.cperiod.currentText())
                             if j == i:
                                 pass
                             elif j > i:
                                 for k in range(i + 1, j + 1):
                                     todo_mths.append(k)
                             else:
                                 for k in range(i + 1, 12):
                                     todo_mths.append(k)
                                 for k in range(j + 1):
                                     todo_mths.append(k)
                         for s in todo_mths:
                             m = 0
                             strt_row.append(0)
                             while m < s:
                                 strt_row[-1] = strt_row[-1] + the_days[m] * self.interval
                                 m += 1
                             strt_row[-1] = strt_row[-1] + self.toprow[1]
                             todo_rows.append(the_days[s] * self.interval)
            load = []
            tgt_col = -1
            data = []
            labels = []
            colors = []
            if self.plottype.currentText() == 'Pie Chart':
                fig = plt.figure(figname, constrained_layout=self.constrained_layout)
                plt.title(titl, fontdict=self.fontprops['Title'])
                if suptitle != '':
                    fig.suptitle(suptitle, fontproperties=self.fontprops['Header'])
                pi2 = plt.subplot(111)
                if self.legend_side == 'Right':
                    plt.subplots_adjust(left=0.1, bottom=0.1, right=0.75)
                elif self.legend_side == 'Left':
                    plt.subplots_adjust(left=0.25, bottom=0.1, right=0.9)
                    loc = 'lower left'
                if self.pie_group is not None:
                    groups = []
                    orders = [[], []]
                    cols = []
                    for c in range(self.order.count()):
                        for itm in self.pie_group:
                            if self.order.item(c).text().lower().find(itm) > -1:
                                orders[0].append(c)
                                break
                        else:
                            orders[1].append(c)
                    strt = 0
                    for c in orders:
                        cols.extend(c)
                        groups.append(list(range(strt, strt + len(c))))
                        strt += len(c)
                else:
                    groups = None
                    cols = list(range(self.order.count()))
                for c in cols:
                    col = self.order.item(c).text()
                    for c2 in range(2, ws.ncols):
                        try:
                            column = ws.cell_value(self.toprow[0], c2).replace('\n',' ')
                        except:
                            column = str(ws.cell_value(self.toprow[0], c2))
                        if self.zone_row > 0 and ws.cell_value(self.zone_row, c2) != '' and ws.cell_value(self.zone_row, c2) is not None:
                            column = ws.cell_value(self.zone_row, c2).replace('\n',' ') + '.' + column
                        if column == col:
                            data.append(0.)
                            labels.append(column)
                            colors.append(self.set_colour(labels[-1]))
                            tot_rows = 0
                            for s in range(len(strt_row)):
                                for row in range(strt_row[s] + 1, strt_row[s] + todo_rows[s] + 1):
                                    try:
                                        data[-1] = data[-1] + ws.cell_value(row, c2)
                                    except:
                                        break # part period
                            break
                tot = sum(data)
                if self.pie_group is not None:
                    re_pct = 0
                    for c in groups[0]:
                        re_pct += data[c]
                    try:
                        re_pct = re_pct * 100. / tot
                    except:
                        pass
                # https://stackoverflow.com/questions/3942878/how-to-decide-font-color-in-white-or-black-depending-on-background-color
                whites = []
                threshold = sqrt(1.05 * 0.05) - 0.05
                if self.percent_on_pie:
                    for c in range(len(colors)):
                        intensity = 0.
                        for i in range(1, 5, 2):
                            colr = int(colors[c][i : i + 2], 16) / 255.0
                            if colr <= 0.04045:
                                colr = colr / 12.92
                            else:
                                colr = pow((colr + 0.055) / 1.055, 2.4)
                            if i == 1: #red
                                intensity = colr * 0.216
                            elif i == 3:
                                intensity = colr * 0.7152
                            else:
                                intensity = colr * 0.0722
                        if intensity < threshold:
                            whites.append(c)
                if self.legend_on_pie: # legend on chart
                    if self.percent_on_pie:
                        patches, texts, autotexts = pi2.pie(data, labels=labels, autopct='%1.1f%%', colors=colors, startangle=90,
                                                            pctdistance=.70)
                        for text in autotexts:
                            text.set_family(self.legend_font.get_family()[0])
                            text.set_style(self.legend_font.get_style())
                            text.set_variant(self.legend_font.get_variant())
                            text.set_stretch(self.legend_font.get_stretch())
                            text.set_weight(self.legend_font.get_weight())
                            text.set_size(self.legend_font.get_size())
                    else:
                        patches, texts = pi2.pie(data, labels=labels, colors=colors, startangle=90,
                                                 pctdistance=.70)
                    for text in texts:
                        text.set_family(self.legend_font.get_family()[0])
                        text.set_style(self.legend_font.get_style())
                        text.set_variant(self.legend_font.get_variant())
                        text.set_stretch(self.legend_font.get_stretch())
                        text.set_weight(self.legend_font.get_weight())
                        text.set_size(self.legend_font.get_size())
                else:
                    loc = 'lower right'
                    if self.percent_on_pie:
                        patches, texts, autotexts = pi2.pie(data, labels=None, autopct='%1.1f%%', colors=colors, startangle=90,
                                                            pctdistance=.70)
                        for text in autotexts:
                            text.set_family(self.legend_font.get_family()[0])
                            text.set_style(self.legend_font.get_style())
                            text.set_variant(self.legend_font.get_variant())
                            text.set_stretch(self.legend_font.get_stretch())
                            text.set_weight(self.legend_font.get_weight())
                            text.set_size(self.legend_font.get_size())
                        for text in texts:
                            text.set_family(self.legend_font.get_family()[0])
                            text.set_style(self.legend_font.get_style())
                            text.set_variant(self.legend_font.get_variant())
                            text.set_stretch(self.legend_font.get_stretch())
                            text.set_weight(self.legend_font.get_weight())
                            text.set_size(self.legend_font.get_size())
                    else:
                        for i in range(len(data)):
                            labels[i] = f'{labels[i]} ({data[i]/tot*100:0.1f}%)'
                        patches, texts = pi2.pie(data, colors=colors, startangle=90)
                    if self.legend_side == 'Right':
                        plt.subplots_adjust(left=0.1, bottom=0.1, right=0.75)
                    elif self.legend_side == 'Left':
                        plt.subplots_adjust(left=0.25, bottom=0.1, right=0.9)
                        loc = 'lower left'
                    fig.legend(patches, labels, loc=loc, ncol=self.legend_ncol, prop=self.legend_font).set_draggable(True)
                if self.pie_group is not None:
                # https://stackoverflow.com/questions/20549016/explode-multiple-slices-of-pie-together-in-matplotlib
                    radfraction = 0.015 # or 0.15
                    for group in groups:
                        ang = np.deg2rad((patches[group[-1]].theta2 + patches[group[0]].theta1) / 2)
                        for j in group:
                            center = radfraction * patches[j].r * np.array([np.cos(ang), np.sin(ang)])
                            patches[j].set_center(center)
                         #   labels[j].set_position(np.array(labels[j].get_position()) + center) Nah
                         #   texts[j].set_position(np.array(texts[j].get_position()) + center)
                for c in whites:
                    autotexts[c].set_color('white')
                p = plt.gcf()
                p.gca().add_artist(plt.Circle((0, 0), 0.40, color='white'))
                if self.pie_group:
                    p.gca().add_artist(plt.text(x=0, y=0, s=f'{re_pct:0.1f}% RE', ha='center', va='center', font=self.legend_font))
                plt.show()
                return
            elif self.plottype.currentText() == 'Heat Map':
                fig = plt.figure(figname)
                if suptitle != '':
                    fig.suptitle(suptitle, fontproperties=self.fontprops['Header'])
                if gridtype != '':
                    plt.grid(axis=gridtype)
                hm2 = plt.subplot(111)
                plt.title(titl, fontdict=self.fontprops['Title'])
                hmdata = []
                days = int(sum(todo_rows) / self.interval)
                for hr in range(self.interval):
                    hmdata.append([])
                    for dy in range(days):
                        hmdata[-1].append(0)
                for c2 in range(2, ws.ncols):
                    try:
                        column = ws.cell_value(self.toprow[0], c2).replace('\n',' ')
                    except:
                        column = str(ws.cell_value(self.toprow[0], c2))
                    if self.zone_row > 0 and ws.cell_value(self.zone_row, c2) != '' and ws.cell_value(self.zone_row, c2) is not None:
                        column = ws.cell_value(self.zone_row, c2).replace('\n',' ') + '.' + column
                    if column == self.target:
                        tgt_col = c2
                for c in range(self.order.count() -1, -1, -1):
                    col = self.order.item(c).text()
                    for c2 in range(2, ws.ncols):
                        try:
                            column = ws.cell_value(self.toprow[0], c2).replace('\n',' ')
                        except:
                            column = str(ws.cell_value(self.toprow[0], c2))
                        if self.zone_row > 0 and ws.cell_value(self.zone_row, c2) != '' and ws.cell_value(self.zone_row, c2) is not None:
                            column = ws.cell_value(self.zone_row, c2).replace('\n',' ') + '.' + column
                        if column == col:
                            dy = 0
                            hr = -1
                            for s in range(len(strt_row)):
                                for row in range(strt_row[s] + 1, strt_row[s] + todo_rows[s] + 1):
                                    hr += 1
                                    if hr > self.interval - 1:
                                        dy += 1
                                        hr = 0
                                    try:
                                        hmdata[hr][dy] += ws.cell_value(row, c2)
                                    except:
                                        pass
                            break
                if tgt_col >= 0:
                    dy = 0
                    hr = -1
                    for s in range(len(strt_row)):
                        for row in range(strt_row[s] + 1, strt_row[s] + todo_rows[s] + 1):
                            hr += 1
                            if hr > self.interval - 1:
                                dy += 1
                                hr = 0
                            try:
                                hmdata[hr][dy] = hmdata[hr][dy] / ws.cell_value(row, tgt_col)
                            except:
                                pass
                miny = 999999
                maxy = 0
                tgap = 0
                for y in range(len(hmdata)):
                    for x in range(len(hmdata[0])):
                        miny = min(miny, hmdata[y][x])
                        maxy = max(maxy, hmdata[y][x])
                if tgt_col >= 0:
                    fmt_str = '{:.3f}'
                    vmax = 1
                    vmin = 0
                    if self.cbar2:
                        tgap = .03
                    if maxy > 1:
                        maxy = 1
                else:
                    fmt_str = '{:,.0f}'
                    if self.cbar2:
                        tgap = (maxy - miny) * .03
                        vmax = maxy
                        vmin = miny
                    else:
                        vmax = None
                        vmin = None
                im2 = hm2.imshow(hmdata, cmap=self.cmap, interpolation='nearest', aspect='auto', vmax=vmax, vmin=vmin)
                yticks = [-0.5]
                if self.interval == 24:
                    for y in range(0, 24, 4):
                        yticks.append(y + 2.5)
                else:
                    for y in range(0, 48, 8):
                        yticks.append(y + 6.5)
                hm2.invert_yaxis()
                hm2.set_yticks(yticks)
                hm2.set_yticklabels(hr_labels, va='bottom', fontdict=self.fontprops['Ticks'])
                if len(hmdata[0]) > 31:
                    day = -0.5
                    xticks = []
                    xticklbls = []
                    for m in todo_mths:
                        xticks.append(day)
                        xticklbls.append(mth_labels[m])
                        day = day + the_days[m]
                    hm2.set_xticks(xticks)
                    hm2.set_xticklabels(xticklbls, ha='left', fontdict=self.fontprops['Ticks'])
                else:
                    curxticks = hm2.get_xticks()
                    xticks = []
                    xticklabels = []
                    for x in range(len(curxticks)):
                        if curxticks[x] >= 0 and curxticks[x] < len(hmdata[0]):
                            xticklabels.append(str(int(curxticks[x] + 1)))
                            xticks.append(curxticks[x])
                    hm2.set_xticks(xticks)
                    hm2.set_xticklabels(xticklabels, fontdict=self.fontprops['Ticks'])
                hm2.set_xlabel('Day', fontdict=self.fontprops['Label'])
                hm2.set_ylabel('Hour', fontdict=self.fontprops['Label'])
                if self.cbar:
                    fig.subplots_adjust(bottom=0.1, right=0.8, top=0.9)
                    cb_ax = fig.add_axes([0.85, 0.1, 0.025, 0.8])
                    cbar = fig.colorbar(im2, cax=cb_ax)
                    if self.cbar2:
                        curcticks = cbar.get_ticks()
                        cticks = [vmin, vmax]
                        if miny != vmin:
                            cticks.append(miny)
                        if maxy != vmax:
                            cticks.append(maxy)
                        for y in range(len(curcticks)):
                            if curcticks[y] < vmin or curcticks[y] > vmax:
                                continue
                            if self.cbar2 and tgap != 0:
                                if miny - tgap < curcticks[y] < miny + tgap:
                                    continue
                                if maxy - tgap < curcticks[y] < maxy + tgap:
                                    continue
                            cticks.append(curcticks[y])
                        cticks.sort()
                        try:
                            cbar.set_ticks(cticks, fontdict=self.fontprops['Ticks'])
                        except:
                            clabels = []
                            for ct in cticks:
                                clabels.append(f'{ct}')
                            cbar.set_ticks(cticks, labels=clabels, fontdict=self.fontprops['Ticks'])
                self.log.setText('Heat map value range: {} to {}'.format(fmt_str, fmt_str).format(miny, maxy))
                QtCore.QCoreApplication.processEvents()
                plt.show()
                return
            overlay_cols = []
            overlay = []
            for o in range(self.no_of_overlays):
                overlay_cols.append([])
                overlay.append([])
            miny = 0
            maxy = 0
            hs = []
            for h in range(self.interval):
                hs.append(h)
            x = hs[:]
            for c2 in range(2, ws.ncols):
                try:
                    column = ws.cell_value(self.toprow[0], c2).replace('\n',' ')
                except:
                    column = str(ws.cell_value(self.toprow[0], c2))
                if self.zone_row > 0 and ws.cell_value(self.zone_row, c2) != '' and ws.cell_value(self.zone_row, c2) is not None:
                    column = ws.cell_value(self.zone_row, c2).replace('\n',' ') + '.' + column
                if column == self.target:
                    tgt_col = c2
                for o in range(self.no_of_overlays):
                    if self.overlay[o] != '<none>':
                        if column[:len(self.overlay[o])] == self.overlay[o] or \
                          ((self.overlay[o] == 'Charge' or self.overlay[o] == 'Underlying Load') and column == self.target):
                            overlay_cols[o].append(c2)
            if len(breakdowns) > 0:
                for c in range(self.order.count() -1, -1, -1):
                    col = self.order.item(c).text()
                    for c2 in range(2, ws.ncols):
                        try:
                            brkdown = ws.cell_value(self.breakdown_row, c2)
                            if brkdown is not None and brkdown != '' and brkdown != breakdowns[0]:
                                continue
                        except:
                            pass
                        try:
                            column = ws.cell_value(self.toprow[0], c2).replace('\n',' ')
                        except:
                            column = str(ws.cell_value(self.toprow[0], c2))
                        if self.zone_row > 0 and ws.cell_value(self.zone_row, c2) != '' and ws.cell_value(self.zone_row, c2) is not None:
                            column = ws.cell_value(self.zone_row, c2).replace('\n',' ') + '.' + column
                        if column == col:
                            data.append([])
                            data[-1] = [0] * len(hs)
                            label.append(column)
                            tot_rows = 0
                            for s in range(len(strt_row)):
                                h = 0
                                for row in range(strt_row[s] + 1, strt_row[s] + todo_rows[s] + 1):
                                    try:
                                        data[-1][h] = data[-1][h] + ws.cell_value(row, c2)
                                    except:
                                        break # part period
                                    h += 1
                                    if h >= self.interval:
                                        h = 0
                                tot_rows += todo_rows[s]
                            for h in range(self.interval):
                                data[-1][h] = data[-1][h] / (tot_rows / self.interval)
                                maxy = max(maxy, data[-1][h])
                                miny = min(miny, data[-1][h])
                            break
                for breakdown in breakdowns[1:]:
                    for c in range(self.order.count() -1, -1, -1):
                        col = self.order.item(c).text()
                        for c2 in range(2, ws.ncols):
                            try:
                                brkdown = ws.cell_value(self.breakdown_row, c2)
                                if brkdown != breakdown:
                                    continue
                            except:
                                continue
                            try:
                                column = ws.cell_value(self.toprow[0], c2).replace('\n',' ')
                            except:
                                column = str(ws.cell_value(self.toprow[0], c2))
                            if self.zone_row > 0 and ws.cell_value(self.zone_row, c2) != '' and ws.cell_value(self.zone_row, c2) is not None:
                                column = ws.cell_value(self.zone_row, c2).replace('\n',' ') + '.' + column
                            if column == col:
                                data.append([])
                                data[-1] = [0] * len(hs)
                                label.append(self.short_legend + column + ' ' + breakdown)
                                tot_rows = 0
                                for s in range(len(strt_row)):
                                    h = 0
                                    for row in range(strt_row[s] + 1, strt_row[s] + todo_rows[s] + 1):
                                        try:
                                            data[-1][h] = data[-1][h] + ws.cell_value(row, c2)
                                        except:
                                            break # part period
                                        h += 1
                                        if h >= self.interval:
                                            h = 0
                                    tot_rows += todo_rows[s]
                                for h in range(self.interval):
                                    data[-1][h] = data[-1][h] / (tot_rows / self.interval)
                                    maxy = max(maxy, data[-1][h])
                                    miny = min(miny, data[-1][h])
                                break
            else:
                for c in range(self.order.count() -1, -1, -1):
                    col = self.order.item(c).text()
                    for c2 in range(2, ws.ncols):
                        try:
                            column = ws.cell_value(self.toprow[0], c2).replace('\n',' ')
                        except:
                            column = str(ws.cell_value(self.toprow[0], c2))
                        if self.zone_row > 0 and ws.cell_value(self.zone_row, c2) != '' and ws.cell_value(self.zone_row, c2) is not None:
                            column = ws.cell_value(self.zone_row, c2).replace('\n',' ') + '.' + column
                        if column == col:
                            data.append([])
                            data[-1] = [0] * len(hs)
                            label.append(column)
                            tot_rows = 0
                            for s in range(len(strt_row)):
                                h = 0
                                for row in range(strt_row[s] + 1, strt_row[s] + todo_rows[s] + 1):
                                    try:
                                        data[-1][h] = data[-1][h] + ws.cell_value(row, c2)
                                    except:
                                        break # part period
                                    h += 1
                                    if h >= self.interval:
                                        h = 0
                                tot_rows += todo_rows[s]
                            for h in range(self.interval):
                                data[-1][h] = data[-1][h] / (tot_rows / self.interval)
                                maxy = max(maxy, data[-1][h])
                                miny = min(miny, data[-1][h])
                            break
            if tgt_col >= 0:
                s = 0
                load = [0] * len(hs)
                tot_rows = 0
                for s in range(len(strt_row)):
                    h = 0
                    for row in range(strt_row[s] + 1, strt_row[s] + todo_rows[s] + 1):
                        try:
                            load[h] = load[h] + ws.cell_value(row, tgt_col)
                        except:
                            print('(3453)', h, row, tgt_col, ws.cell_value(row, tgt_col), strt_row, todo_rows)
                        h += 1
                        if h >= self.interval:
                           h = 0
                    tot_rows += todo_rows[s]
                for h in range(self.interval):
                    load[h] = load[h] / (tot_rows / self.interval)
                    maxy = max(maxy, load[h])
                    miny = min(miny, load[h])
            for o in range(self.no_of_overlays):
                if len(overlay_cols[o]) > 0:
                    if self.overlay[o] == 'Underlying Load':
                        overlay_cols[o] = [overlay_cols[o][-1]]
                    for h in range(len(hs)):
                        overlay[o].append(0)
                    tot_rows = 0
                    for s in range(len(strt_row)):
                        h = 0
                        for row in range(strt_row[s] + 1, strt_row[s] + todo_rows[s] + 1):
                            for col in overlay_cols[o]:
                                try:
                                    overlay[o][h] = overlay[o][h] + ws.cell_value(row, col)
                                except:
                                 #   if row >= ws.nrows:
                                 #       self.log.setText(f'Period may be incomplete (8)')
                                 #       break
                                    self.log.setText(f'Data error with {self.overlay[o]} ({ssCol(col, base=0)}{row + 1}). Period may be incomplete (8)')
                                    return
                            h += 1
                            if h >= self.interval:
                               h = 0
                        tot_rows += todo_rows[s]
                    for h in range(self.interval):
                        overlay[o][h] = overlay[o][h] / (tot_rows / self.interval)
                        # 'overlay_type'
                        if self.overlay_type[o] == 'Average':
                            overlay[o][h] = overlay[o][h] / self.interval
                        maxy = max(maxy, overlay[o][h])
                        miny = min(miny, overlay[o][h])
            loc = 'lower right'
            if self.plottype.currentText() == 'Line Chart' or self.plottype.currentText() == 'Step Line Chart':
                fig = plt.figure(figname + '_' + self.period.currentText().lower(),
                                 constrained_layout=self.constrained_layout)
                if suptitle != '':
                    fig.suptitle(suptitle, fontproperties=self.fontprops['Header'])
                if gridtype != '':
                    plt.grid(axis=gridtype)
                lc2 = plt.subplot(111)
                if self.legend_side == 'Right':
                    plt.subplots_adjust(left=0.1, bottom=0.1, right=0.75)
                elif self.legend_side == 'Left':
                    plt.subplots_adjust(left=0.25, bottom=0.1, right=0.9)
                    loc = 'lower left'
                plt.title(titl, fontdict=self.fontprops['Title'])
                if self.plottype.currentText() == 'Line Chart':
                    for c in range(len(data)):
                        lc2.plot(x, data[c], linewidth=self.linSpin.value(), label=label[c], color=self.set_colour(label[c]))
                    if len(load) > 0:
                        lc2.plot(x, load, linewidth=self.tgtSpin.value(), label=self.short_legend + self.target,
                                 color=self.set_colour(self.target), linestyle=self.tgtLine.currentText())
                    for o in range(self.no_of_overlays):
                        if len(overlay[o]) > 0:
                            lc2.plot(x, overlay[o], linewidth=self.ovrSpin[o].value(), label=self.short_legend + self.overlay[o],
                                     color=self.overlay_colour[o], linestyle=self.ovrLine[o].currentText())
                else:
                    for c in range(len(data)):
                        lc2.step(x, data[c], linewidth=self.linSpin.value(), label=label[c], color=self.set_colour(label[c]),
                                 linestyle=self.linLine.currentText())
                    if len(load) > 0:
                        lc2.step(x, load, linewidth=self.tgtSpin.value(), label=self.short_legend + self.target,
                                 color=self.set_colour(self.target), linestyle=self.tgtLine.currentText())
                    for o in range(self.no_of_overlays):
                        if len(overlay[o]) > 0:
                            lc2.step(x, overlay[o], linewidth=self.ovrSpin[o].value(), label=self.short_legend + self.overlay[o],
                                     color=self.overlay_colour[o], linestyle=self.ovrLine[o].currentText())
                yminmax = set_ylimits(miny, maxy)
                lc2.legend(bbox_to_anchor=[0.5, -0.1], loc='center', ncol=(len(data) + 1),
                              prop=self.legend_font)
                plt.ylim(yminmax)
                plt.xlim([0, self.interval - 1])
                lc2.set_xticks(ticks)
                lc2.set_xticklabels(hr_labels, fontdict=self.fontprops['Ticks'])
                lc2.set_xlabel('Hour of the Day', fontdict=self.fontprops['Label'])
                lc2.set_ylabel('Power (MW/MWh)', fontdict=self.fontprops['Label'])
                lc2.tick_params(colors=self.fontprops['Ticks']['color'], which='both')
                if self.y2_label != '':
                    secax = lc2.secondary_yaxis(location='right', functions=(self.transform_y2, self.inverse_y2))
                    secax.set_ylabel(self.y2_label)
                zp = ZoomPanX()
                f = zp.zoom_pan(lc2, base_scale=1.2) # enable scrollable zoom
                plt.show()
                del zp
            elif self.plottype.currentText() in ['Cumulative', 'Step Chart']:
                if self.plottype.currentText() == 'Cumulative':
                    step = None
                else:
                    step = 'pre' # 'post'
                fig = plt.figure(figname + '_' + self.period.currentText().lower(),
                                 constrained_layout=self.constrained_layout)
                if suptitle != '':
                    fig.suptitle(suptitle, fontproperties=self.fontprops['Header'])
                if gridtype != '':
                    plt.grid(axis=gridtype)
                cu2 = plt.subplot(111)
                if self.legend_side == 'Right':
                    plt.subplots_adjust(left=0.1, bottom=0.1, right=0.75)
                elif self.legend_side == 'Left':
                    plt.subplots_adjust(left=0.25, bottom=0.1, right=0.9)
                    loc = 'lower left'
                plt.title(titl, fontdict=self.fontprops['Title'])
                if self.percentage.isChecked():
                    totals = [0.] * len(x)
                    bottoms = [0.] * len(x)
                    values = [0.] * len(x)
                    for c in range(len(data)):
                        for h in range(len(data[c])):
                            totals[h] = totals[h] + data[c][h]
                    for h in range(len(data[0])):
                        values[h] = data[0][h] / totals[h] * 100.
                    cu2.fill_between(x, 0, values, label=label[0], color=self.set_colour(label[0]), hatch=self.set_hatch(label[0]),
                                     alpha=self.alpha_fill, step=step)
                    for c in range(1, len(data)):
                        for h in range(len(data[c])):
                            bottoms[h] = values[h]
                            values[h] = values[h] + data[c][h] / totals[h] * 100.
                        cu2.fill_between(x, bottoms, values, label=label[c], color=self.set_colour(label[c]), hatch=self.set_hatch(label[c]),
                                         alpha=self.alpha_fill, step=step)
                    maxy = 100
                    cu2.set_ylabel('Power (%)', fontdict=self.fontprops['Label'])
                else:
                    if self.target == '<none>':
                        cu2.fill_between(x, 0, data[0], label=label[0], color=self.set_colour(label[0]), hatch=self.set_hatch(label[0]),
                                         alpha=self.alpha_fill, step=step)
                        for c in range(1, len(data)):
                            for h in range(len(data[c])):
                                data[c][h] = data[c][h] + data[c - 1][h]
                                maxy = max(maxy, data[c][h])
                            cu2.fill_between(x, data[c - 1], data[c], label=label[c], color=self.set_colour(label[c]), hatch=self.set_hatch(label[c]),
                                             alpha=self.alpha_fill, step=step)
                    else:
                #        pattern = ['-', '+', 'x', '\\', '*', 'o', 'O', '.']
                #        pat = 0
                        full = []
                        for h in range(len(load)):
                           full.append(min(load[h], data[0][h]))
                        cu2.fill_between(x, 0, full, label=label[0], color=self.set_colour(label[0]), hatch=self.set_hatch(label[0]),
                                         alpha=self.alpha_fill, step=step)
                        for h in range(len(full)):
                            if data[0][h] > full[h]:
                                if self.spill_label.text() != '':
                                    cu2.fill_between(x, full, data[0], alpha=self.alpha, color=self.set_colour(label[0]), hatch=self.set_hatch(label[0]),
                                        label=label[0] + ' ' + self.spill_label.text(), step=step)
                                else:
                                    cu2.fill_between(x, full, data[c], alpha=self.alpha, color=self.set_colour(label[c]), hatch=self.set_hatch(label[c]),
                                                     step=step)
                                break
                        for c in range(1, len(data)):
                            full = []
                            for h in range(len(data[c])):
                                data[c][h] = data[c][h] + data[c - 1][h]
                                maxy = max(maxy, data[c][h])
                                full.append(max(min(load[h], data[c][h]), data[c - 1][h]))
                            cu2.fill_between(x, data[c - 1], full, label=label[c], color=self.set_colour(label[c]), hatch=self.set_hatch(label[c]),
                                             alpha=self.alpha_fill, step=step)
                 #           pat += 1
                 #           if pat >= len(pattern):
                 #               pat = 0
                            for h in range(len(full)):
                                if data[c][h] > full[h] + self.margin_of_error:
                                    if self.spill_label.text() != '':
                                        cu2.fill_between(x, full, data[c], alpha=self.alpha, color=self.set_colour(label[c]), hatch=self.set_hatch(label[c]),
                                                        label=label[c] + ' ' + self.spill_label.text(), step=step)
                                    else:
                                        cu2.fill_between(x, full, data[c], alpha=self.alpha, color=self.set_colour(label[c]), hatch=self.set_hatch(label[c]),
                                                        step=step)
                                    break
                        top = data[0][:]
                        for d in range(1, len(data)):
                            for h in range(len(top)):
                                top[h] = max(top[h], data[d][h])
                        if self.plottype.currentText() == 'Cumulative':
                            if self.alpha == 0:
                                cu2.plot(x, top, color='gray', linestyle='dashed')
                            else:
                                cu2.plot(x, top, color='gray')
                        else:
                            if self.alpha == 0:
                                cu2.step(x, top, color='gray', linestyle='dashed')
                            else:
                                cu2.step(x, top, color='gray')
                        short = []
                        do_short = False
                        for h in range(len(load)):
                            if load[h] > data[c][h] + self.margin_of_error:
                                do_short = True
                            short.append(max(data[c][h], load[h]))
                        if do_short:
                            if self.colours['shortfall'][0].lower() != 'w' and self.colours['shortfall'].lower() != '#ffffff':
                                cu2.fill_between(x, data[c], short, label='Shortfall', color=self.set_colour('shortfall'),
                                                 alpha=self.alpha_fill, step=step)
                        if self.plottype.currentText() == 'Cumulative':
                            cu2.plot(x, load, linewidth=self.tgtSpin.value(), label=self.short_legend + self.target, color=self.set_colour(self.target),
                                     linestyle=self.tgtLine.currentText())
                        else:
                            cu2.step(x, load, linewidth=self.tgtSpin.value(), label=self.short_legend + self.target, color=self.set_colour(self.target),
                                     linestyle=self.tgtLine.currentText())
                    for o in range(self.no_of_overlays):
                        if len(overlay[o]) > 0:
                            if self.plottype.currentText() == 'Cumulative':
                                cu2.plot(x, overlay[o], linewidth=self.ovrSpin[o].value(), label=self.short_legend + self.overlay[o],
                                         color=self.overlay_colour[o], linestyle=self.ovrLine[o].currentText())
                            else:
                                cu2.step(x, overlay[o], linewidth=self.ovrSpin[o].value(), label=self.short_legend + self.overlay[o],
                                         color=self.overlay_colour[o], linestyle=self.ovrLine[o].currentText())
                    cu2.set_ylabel('Power (MW/MWh)', fontdict=self.fontprops['Label'])
         #       miny = 0
                yminmax = set_ylimits(miny, maxy)
                cu2.legend(bbox_to_anchor=[0.5, -0.1], loc='center', ncol=(len(data) + 2), prop=self.legend_font)
                plt.ylim(yminmax)
                plt.xlim([0, self.interval - 1])
                cu2.set_xticks(ticks)
                cu2.set_xticklabels(hr_labels, fontdict=self.fontprops['Ticks'])
                cu2.set_xlabel('Hour of the Day', fontdict=self.fontprops['Label'])
                cu2.tick_params(colors=self.fontprops['Ticks']['color'], which='both')
                if self.y2_label != '':
                    secax = cu2.secondary_yaxis(location='right', functions=(self.transform_y2, self.inverse_y2))
                    secax.set_ylabel(self.y2_label)
                zp = ZoomPanX()
                f = zp.zoom_pan(cu2, base_scale=1.2) # enable scrollable zoom
                plt.show()
                del zp
            elif self.plottype.currentText() == 'Bar Chart':
                fig = plt.figure(figname + '_' + self.period.currentText().lower(),
                                 constrained_layout=self.constrained_layout)
                if suptitle != '':
                    fig.suptitle(suptitle, fontproperties=self.fontprops['Header'])
                if gridtype != '':
                    plt.grid(axis=gridtype)
                bc2 = plt.subplot(111)
                if self.legend_side == 'Right':
                    plt.subplots_adjust(left=0.1, bottom=0.1, right=0.75)
                elif self.legend_side == 'Left':
                    plt.subplots_adjust(left=0.25, bottom=0.1, right=0.9)
                    loc = 'lower left'
                plt.title(titl, fontdict=self.fontprops['Title'])
                if self.percentage.isChecked():
                    miny = 0
                    totals = [0.] * len(x)
                    bottoms = [0.] * len(x)
                    values = [0.] * len(x)
                    for c in range(len(data)):
                        for h in range(len(data[c])):
                            totals[h] = totals[h] + data[c][h]
                    for h in range(len(data[0])):
                        values[h] = data[0][h] / totals[h] * 100.
                    bc2.bar(x, values, label=label[0], color=self.set_colour(label[0]), alpha=self.alpha_fill,
                            hatch=self.set_hatch(label[0]))
                    for c in range(1, len(data)):
                        for h in range(len(data[c])):
                            bottoms[h] = bottoms[h] + values[h]
                            values[h] = data[c][h] / totals[h] * 100.
                        bc2.bar(x, values, bottom=bottoms, label=label[c], color=self.set_colour(label[c]), alpha=self.alpha_fill,
                                hatch=self.set_hatch(label[c]))
                    maxy = 100
                    bc2.set_ylabel('Power (%)', fontdict=self.fontprops['Label'])
                else:
                    if self.target == '<none>':
                        bc2.bar(x, data[0], label=label[0], color=self.set_colour(label[0]), alpha=self.alpha_fill,
                                hatch=self.set_hatch(label[0]))
                        bottoms = [0.] * len(x)
                        for c in range(1, len(data)):
                            for h in range(len(data[c])):
                                bottoms[h] = bottoms[h] + data[c - 1][h]
                                maxy = max(maxy, data[c][h] + bottoms[h])
                            bc2.bar(x, data[c], bottom=bottoms, label=label[c], color=self.set_colour(label[c]), alpha=self.alpha_fill,
                                    hatch=self.set_hatch(label[c]))
                    else:
                        bc2.bar(x, data[0], label=label[0], color=self.set_colour(label[0]), alpha=self.alpha_fill,
                                hatch=self.set_hatch(label[0]))
                        bottoms = [0.] * len(x)
                        for c in range(1, len(data)):
                            for h in range(len(data[c])):
                                bottoms[h] = bottoms[h] + data[c - 1][h]
                                maxy = max(maxy, data[c][h] + bottoms[h])
                            bc2.bar(x, data[c], bottom=bottoms, label=label[c], color=self.set_colour(label[c]), alpha=self.alpha_fill,
                                    hatch=self.set_hatch(label[c]))
                        bc2.plot(x, load, linewidth=self.tgtSpin.value(), label=self.short_legend + self.target, color=self.set_colour(self.target),
                                 linestyle=self.tgtLine.currentText())
                    for o in range(self.no_of_overlays):
                        if len(overlay[o]) > 0:
                            bc2.plot(x, overlay[o], linewidth=self.ovrSpin[o].value(), label=self.short_legend + self.overlay[o],
                                     color=self.overlay_colour[o], linestyle=self.ovrLine[o].currentText())
                    bc2.set_ylabel('Power (MW/MWh)', fontdict=self.fontprops['Label'])
        #        miny = 0
                yminmax = set_ylimits(miny, maxy)
                if self.legend_side == 'Right':
                    plt.subplots_adjust(left=0.1, bottom=0.1, right=0.75)
                elif self.legend_side == 'Left':
                    plt.subplots_adjust(left=0.25, bottom=0.1, right=0.9)
                    loc = 'lower left'
                bc2.legend(bbox_to_anchor=[0.5, -0.1], loc='center', ncol=(len(data) + 2), prop=self.legend_font)
                plt.ylim(yminmax)
                plt.xlim([0, self.interval - 1])
                bc2.set_xticks(ticks)
                bc2.set_xticklabels(hr_labels, fontdict=self.fontprops['Ticks'])
                bc2.set_xlabel('Hour of the Day', fontdict=self.fontprops['Label'])
                bc2.tick_params(colors=self.fontprops['Ticks']['color'], which='both')
                if self.y2_label != '':
                    secax = bc2.secondary_yaxis(location='right', functions=(self.transform_y2, self.inverse_y2))
                    secax.set_ylabel(self.y2_label)
                zp = ZoomPanX()
                f = zp.zoom_pan(bc2, base_scale=1.2) # enable scrollable zoom
                plt.show()
                del zp

    def corrClicked(self):
        if self.show_contribution:
            corrtxt = ['month', 'contribution']
            sortby = ''
        else:
            corrtxt = ['coefficient', 'correlation']
            sortby = corrtxt[0]
        if self.book is None:
            self.log.setText('Error accessing Workbook.')
            return
        if self.order.count() == 0:
            self.log.setText('No columns chosen.')
            return
        isheet = self.sheet.currentText()
        if isheet == '':
            self.log.setText('Sheet not set.')
            return
        if self.target == '<none>':
            self.log.setText('Target not set.')
            return
        ws = self.book.sheet_by_name(isheet)
        if self.toprow is None:
            tech_row = -1
            row = 0
            self.zone_row =  -1
            self.breakdown_row = -1
            while row < ws.nrows:
                if ws.cell_value(row, 0) in ['Split', 'Breakdown']:
                    self.breakdown_row = row
                elif ws.cell_value(row, 0) == 'Technology':
                    tech_row = row
                elif ws.cell_value(row, 0) == 'Zone':
                    self.zone_row = row
                elif ws.cell_value(row, 0) in ['Hour', 'Interval', 'Trading Interval']:
                    if ws.cell_value(row, 1) != 'Period':
                        self.log.setText(isheet + ' sheet format incorrect')
                        return
                    if tech_row >= 0:
                        self.toprow = [tech_row, row]
                    else:
                        self.toprow = [row, row]
                    self.rows = ws.nrows - (row + 1)
                    break
        try:
            year = int(ws.cell_value(self.toprow[1] + 1, 1)[:4])
            if year % 4 == 0 and year % 100 != 0 or year % 400 == 0:
                self.leapyear = True
            else:
                self.leapyear = False
        except:
            self.leapyear = False
            year = ''
        the_days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
        if self.leapyear: #rows == 8784: # leap year
            the_days[1] = 29
        mth_labels = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        periods = []
        objects = []
        per1 = self.period.currentText()
        per2 = self.cperiod.currentText()
        if self.show_contribution and per1 in  ['<none>', 'Year']:
            per1 = 'Jan'
            per2 = 'Dec'
        if per1 in ['<none>', 'Year']: # full year of hourly figures
            strt_row = [self.toprow[1]]
            todo_rows = [self.rows]
            per2 = ''
            if self.show_contribution:
                periods = mth_labels[:]
        else:
            strt_row = []
            todo_rows = []
            if per1 in self.seasons.keys():
                strt_row = []
                todo_rows = []
                for s in self.seasons[per1]:
                    m = 0
                    strt_row.append(0)
                    while m < s:
                        strt_row[-1] = strt_row[-1] + the_days[m] * self.interval
                        m += 1
                    strt_row[-1] = strt_row[-1] + self.toprow[1]
                    todo_rows.append(the_days[s] * self.interval)
                    if self.show_contribution:
                        periods.append(mth_labels[m])
                per2 = ''
            else:
                i = mth_labels.index(per1)
                todo_mths = [i]
                if per2 != '<none>':
                    j = mth_labels.index(per2)
                    if j == i:
                        pass
                    elif j > i:
                        for k in range(i + 1, j + 1):
                            todo_mths.append(k)
                    else:
                        for k in range(i + 1, 12):
                            todo_mths.append(k)
                        for k in range(j + 1):
                            todo_mths.append(k)
                for s in todo_mths:
                    m = 0
                    strt_row.append(0)
                    while m < s:
                        strt_row[-1] = strt_row[-1] + the_days[m] * self.interval
                        m += 1
                    strt_row[-1] = strt_row[-1] + self.toprow[1]
                    todo_rows.append(the_days[s] * self.interval)
                    if self.show_contribution:
                        periods.append(mth_labels[m])
        tgt = []
        tgt_col = -1
        data = []
        best_corr = 0
        best_fac = ''
        total_tgt = 0
        total_data = []
        for c2 in range(2, ws.ncols):
            try:
                column = ws.cell_value(self.toprow[0], c2).replace('\n',' ')
            except:
                column = str(ws.cell_value(self.toprow[0], c2))
            if self.zone_row > 0 and ws.cell_value(self.zone_row, c2) != '' and ws.cell_value(self.zone_row, c2) is not None:
                column = ws.cell_value(self.zone_row, c2).replace('\n',' ') + '.' + column
            if column == self.target:
                tgt_col = c2
        if tgt_col < 0:
            self.log.setText('Target not found.')
            return
        base_tgt = []
        for row in range(self.toprow[1] + 1, self.toprow[1] + self.rows + 1):
            base_tgt.append(ws.cell_value(row, tgt_col))
        table = []
        if self.show_contribution:
            fields = ['month']
            decpts = [0]
        else:
            fields = ['aggregation', 'object'] + corrtxt
            decpts = [0, 0, 4]
        rows = []
        if self.rows > 8784:
            mult = 2
        else:
            mult = 1
        if self.show_contribution:
            aggregates = {'Month': 24 * 31 * mult}
        else:
            aggregates = {'Hour': 1, 'Day': 24 * mult, 'Week': 24 * 7 * mult,
                          'Fortnight': 24 * 14 * mult, 'Month': 24 * 31 * mult}
        debug = False
        if debug:
            cf = open('debug_corr.csv', 'w')
        for key, agg in aggregates.items():
            tgt = []
            tagg = 0
            atgt = 0
            for s in range(len(strt_row)):
                strt = strt_row[s] - self.toprow[1]
                for row in range(strt, strt + todo_rows[s]):
                    atgt += base_tgt[row]
                    tagg += 1
                    if tagg == agg:
                        tgt.append(atgt) # / tagg)
                        tagg = 0
                        atgt = 0
                if tagg > 0:
                    tgt.append(atgt) # / tagg)
                    tagg = 0
                    atgt = 0
            if not self.show_contribution and len(tgt) < 2:
                continue
            data = []
            for c in range(self.order.count() -1, -1, -1):
                col = self.order.item(c).text()
                for c2 in range(2, ws.ncols):
                    try:
                        column = ws.cell_value(self.toprow[0], c2).replace('\n',' ')
                    except:
                        column = str(ws.cell_value(self.toprow[0], c2))
                    if self.zone_row > 0 and ws.cell_value(self.zone_row, c2) != '' and ws.cell_value(self.zone_row, c2) is not None:
                        column = ws.cell_value(self.zone_row, c2).replace('\n',' ') + '.' + column
                    if column == col:
                        data.append([])
                        total_data.append(0.)
                        sagg = 0
                        asrc = 0
                        for s in range(len(strt_row)):
                            for row in range(strt_row[s] + 1, strt_row[s] + todo_rows[s] + 1):
                                asrc += ws.cell_value(row, c2)
                                sagg += 1
                                if sagg == agg:
                                    data[-1].append(asrc) # / tagg)
                                    sagg = 0
                                    asrc = 0
                            if sagg > 0:
                                data[-1].append(asrc) # / tagg)
                                sagg = 0
                                asrc = 0
                        if self.show_contribution:
                            fields.append(col.lower())
                            decpts.append(4)
                            continue
                            tgt_total = 0
                            src_total = 0
                            for value in tgt:
                                tgt_total += value
                            for value in data[-1]:
                                src_total += value
                            try:
                                corr = src_total / tgt_total
                            except:
                                corr = 0
                        else:
                            try:
                                corr = np.corrcoef(tgt, data[-1])
                                if np.isnan(corr.item((0, 1))):
                                    cor = 'n/a'
                                    corr = 0
                            except:
                                cor = 'n/a'
                                corr = 0
                            else:
                                corr = corr.item((0, 1))
                        if corr == 0:
                            cor = 'n/a'
                        elif abs(corr) < 0.1:
                            cor = 'None'
                        elif abs(corr) < 0.3:
                            cor = 'Little if any'
                        elif abs(corr) < 0.5:
                            cor = 'Low'
                        elif abs(corr) < 0.7:
                            cor = 'Moderate'
                        elif abs(corr) < 0.9:
                            cor = 'High'
                        else:
                            cor = 'Very high'
                        if abs(corr) > best_corr:
                            best_corr = corr
                            best_fac = key + ': ' + self.order.item(c).text()
                        if self.show_contribution:
                            rows.append([self.order.item(c).text(), corr, cor])
                        else:
                            rows.append([key, self.order.item(c).text(), corr, cor])
                       #     break
            if debug:
                for i in range(len(tgt)):
                    line = '%s,%d,%0.2f,' % (key, agg, tgt[i])
                    for d in range(len(data)):
                        line += '%0.2f,' % data[d][i]
                    cf.write(line + '\n')
        if self.show_contribution:
            for m in range(len(periods)):
                if debug:
                    line = periods[m] + ',' + str(tgt[m]) + ','
                total_tgt += tgt[m]
                rows.append([periods[m]])
                for o in range(len(fields) - 1):
                    if debug:
                        line += str(data[o][m]) + ','
                    total_data[o] += data[o][m]
                    pct = '{:.1f}%'.format(data[o][m] * 100. / tgt[m])
                    rows[-1].append(pct)
                if debug:
                    cf.write(line + '\n')
            rows.append(['Total'])
            for o in range(len(fields) - 1):
                pct = '{:.1f}%'.format(total_data[o] * 100. / total_tgt)
                rows[-1].append(pct)
        if debug:
            cf.close()
        title = '%s %s %s' % (corrtxt[1].title(), self.target, per1)
        if per2 != '' and per2 != '<none>':
            title += ' to ' + per2
        dialog = Table(rows, fields=fields, title=title, decpts=decpts, sortby=sortby, reverse=True, save_folder=self.scenarios)
        dialog.exec()
        self.log.setText('Best %s: %s %.4f' % (corrtxt[1], best_fac, best_corr))


if "__main__" == __name__:
    app = QtWidgets.QApplication(sys.argv)
    ex = PowerPlot()
    app.exec()
    app.deleteLater()
    sys.exit()
