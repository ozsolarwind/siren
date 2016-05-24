#!/usr/bin/python
#
#  Copyright (C) 2015-2016 Sustainable Energy Now Inc., Angus King
#
#  displaytable.py - This file is part of SIREN.
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

import csv
import os
import sys
import types
import xlwt
from PyQt4 import QtCore
from PyQt4 import QtGui

import displayobject
from sirenicons import Icons


class FakeObject:
    def __init__(self, fake_object, fields):
        f = -1
        for i in range(len(fake_object)):
            if isinstance(fake_object[i], list):
                for j in range(len(fake_object[i])):
                    f += 1
                    setattr(self, fields[f], fake_object[i][j])
            else:
                f += 1
                setattr(self, fields[f], fake_object[i])


class Table(QtGui.QDialog):
    def __init__(self, objects, parent=None, fields=None, fossil=True, sumby=None, sumfields=None, units='', title=None, \
                 save_folder='', edit=False, sortby=None):
        super(Table, self).__init__(parent)
        if isinstance(objects, list) and isinstance(objects[0], list):
            fakes = []
            for row in objects:
                fakes.append(FakeObject(row, fields))
            self.objects = fakes
        elif isinstance(objects, dict):
            fakes = []
            for key, value in objects.iteritems():
                fakes.append(FakeObject([key, value], fields))
            self.objects = fakes
        else:
            self.objects = objects
        self.icons = Icons()
        self.fields = fields
        self.fossil = fossil
#       somewhere we need to cater for no fields and a sumby
        self.sumby = sumby
        self.sumfields = sumfields
        self.units = units
        self.title = title
        self.edit_table = edit
        self.recur = False
        self.replaced = None
        if self.edit_table:
            self.title_word = ['Edit', 'Export']
        else:
            self.title_word = ['Display', 'Save']
        self.save_folder = save_folder
        if self.sumfields is not None:
            if isinstance(self.sumfields, str):
                self.sumfields = [self.sumfields]
        if sumby is not None:
            if isinstance(self.sumby, str):
                self.sumby = [self.sumby]
            if len(self.sumby) == 1 and self.sumfields is None:
                self.sumfields = self.sumby
            for i in range(len(self.sumfields)):
                for j in range(len(self.fields)):
                    if self.sumfields[i] == self.fields[j]:
                        self.fields.insert(j + 1, '%' + str(i))
                        break
        if self.title is None:
            self.setWindowTitle('SIREN - ' + self.title_word[0] + ' ' + getattr(objects[0], '__module__') + 's')
        else:
            self.setWindowTitle('SIREN - ' + self.title_word[0] + ' ' + self.title)
        msg = '(Right click column header to sort)'
        try:
            if getattr(objects[0], '__module__') == 'Station':
                msg = '(Left or right click row to display, ' + msg[1:]
        except:
            pass
        buttonLayout = QtGui.QHBoxLayout()
        self.message = QtGui.QLabel(msg)
        self.quitButton = QtGui.QPushButton(self.tr('&Quit'))
        buttonLayout.addWidget(self.quitButton)
        self.connect(self.quitButton, QtCore.SIGNAL('clicked()'), \
                    self.quit)
        if self.edit_table:
            self.replaceButton = QtGui.QPushButton(self.tr('Save'))
            buttonLayout.addWidget(self.replaceButton)
            self.connect(self.replaceButton, QtCore.SIGNAL('clicked()'), \
                        self.replacetbl)
        self.saveButton = QtGui.QPushButton(self.tr(self.title_word[1]))
        buttonLayout.addWidget(self.saveButton)
        self.connect(self.saveButton, QtCore.SIGNAL('clicked()'), \
                    self.saveit)
        buttons = QtGui.QFrame()
        buttons.setLayout(buttonLayout)
        layout = QtGui.QGridLayout()
        self.table = QtGui.QTableWidget()
        self.populate()
        if self.sumfields is None:
            self.table.setRowCount(len(self.entry))
        else:
            self.table.setRowCount(len(self.entry) + len(self.sums))
        self.table.setColumnCount(len(self.labels))
        if self.fields is None:
            labels = sorted(map(lambda x: self.nice(x), self.labels.keys()))
        else:
            labels = map(lambda x: self.nice(x), self.fields)
        self.table.setHorizontalHeaderLabels(labels)
        for cl in range(self.table.columnCount()):
            self.table.horizontalHeaderItem(cl).setIcon(QtGui.QIcon('blank.png'))
        self.headers = self.table.horizontalHeader()
        self.headers.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.headers.customContextMenuRequested.connect(self.header_click)
        self.headers.setSelectionMode(QtGui.QAbstractItemView.SingleSelection)
        self.table.verticalHeader().setVisible(False)
      #   self.table.setEditTriggers(QtGui.QAbstractItemView.NoEditTriggers)
        if self.edit_table:
            self.table.setEditTriggers(QtGui.QAbstractItemView.CurrentChanged)
        else:
            self.table.setEditTriggers(QtGui.QAbstractItemView.SelectedClicked)
            self.table.setSelectionBehavior(QtGui.QAbstractItemView.SelectRows)
        layout.addWidget(self.table, 0, 0)
        layout.addWidget(self.message, 1, 0)
        layout.addWidget(buttons, 2, 0)
        self.sort_col = 0
        self.sort_asc = False
        if sortby is None:
            self.order(0)
        else:
            self.order(self.fields.index(sortby))
        if self.sumfields is not None:
            fmat_str = []
            clv = []
            clp = []
            totl = []
            for i in range(len(self.sumfields)):
                clp.append(-1)
            for f in range(len(self.sumfields)):
                fmat_str.append('{:,.' + str(self.lens[self.sumfields[f]][1]) + 'f}')
                clv.append(self.fields.index(self.sumfields[f]))
            if self.sumby is not None:
                clk = self.fields.index(self.sumby[0])
                for i in range(len(self.sumfields)):
                    clp[i] = clv[i] + 1
                    totl.append(self.sums['~~'][i])
            rw = len(self.entry) - 1
            for key, value in iter(sorted(self.sums.iteritems())):
                if key == '~~':
                    tkey = 'Total'
                else:
                    tkey = key
                rw += 1
                if self.sumby is not None:
                    if self.sumby[0] == 'technology' and tkey != 'Total':
                        icon = self.icons.getIcon(tkey)
                        icon_item = QtGui.QTableWidgetItem(str(tkey))
                        icon_item.setIcon(QtGui.QIcon(icon))
                        self.table.setItem(rw, clk, QtGui.QTableWidgetItem(icon_item))
                    else:
                        self.table.setItem(rw, clk, QtGui.QTableWidgetItem(str(tkey)))
                        if not isinstance(tkey, str):
                            self.table.item(rw, clk).setTextAlignment(130)  # x'82'
                for f in range(len(self.sumfields)):
                    self.table.setItem(rw, clv[f], QtGui.QTableWidgetItem(fmat_str[f].format(value[f])))
                    self.table.item(rw, clv[f]).setTextAlignment(130)  # x'82'
                    if clp[f] > 0 and totl[f] > 0:
                        self.table.setItem(rw, clp[f], QtGui.QTableWidgetItem('{:.1%}'.format(float(value[f]) / \
                          float(totl[f])) + ' '))
                        self.table.item(rw, clp[f]).setTextAlignment(130)  # x'82'
        self.table.resizeColumnsToContents()
        width = 0
        for cl in range(self.table.columnCount()):
            width += self.table.columnWidth(cl)
        width += 50
        height = self.table.rowHeight(1) * (self.table.rowCount() + 3)
        screen = QtGui.QDesktopWidget().availableGeometry()
        if height > (screen.height() - 70):
            height = screen.height() - 70
        self.setLayout(layout)
        size = QtCore.QVariant(QtCore.QSize(int(width), int(height))).toSize()
        self.resize(size)
        self.updated = QtCore.pyqtSignal(QtGui.QLabel)   # ??
        QtGui.QShortcut(QtGui.QKeySequence('q'), self, self.quit)
        if self.edit_table:
            self.table.cellChanged.connect(self.item_changed)
        else:
            self.table.cellClicked.connect(self.item_selected)

    def nice(self, string):
        if string == '':
            strout = string
        else:
            strout = string.replace('_', ' ').title()
            strout = strout.replace('Cf', 'CF')
            strout = strout.replace('Lcoe', 'LCOE')
            strout = strout.replace('Npv', 'NPV')
            strout = strout.replace('Pv', 'PV')
            if string != '' and string in self.units:
                i = self.units.find(string)
                j = self.units.find(' ', i)
                if j < 0:
                    j = len(self.units)
                strout = strout + ' (' + self.units[i + 1 + len(string):j] + ')'
        self.hdrs[strout] = string
        return strout

    def populate(self):
        self.labels = {}
        self.lens = {}
        self.hdrs = {}
        if self.fields is not None:
            if '#' in self.fields:
                self.labels['#'] = 'int'
                self.lens['#'] = [len(str(len(self.objects))), 0]   # or int(math.log10(...))+1
            for fld in self.fields:
                if fld[0] == '%':
                    self.labels[fld] = 'str'
                    self.lens[fld] = 0
        for thing in self.objects:
            for prop in dir(thing):
                if prop[:2] != '__' and prop[-2:] != '__':
                    if self.fields is not None:
                        if prop not in self.fields:
                            continue
                    attr = getattr(thing, prop)
                    if attr is None:
                        continue
                    if prop not in self.labels:
                        if type(attr) is types.IntType:
                            self.labels[prop] = 'int'
                        elif type(attr) is types.FloatType:
                            self.labels[prop] = 'float'
                        else:
                            self.labels[prop] = 'str'
                    if type(attr) is types.IntType:
                        if self.labels[prop] == 'str':
                            self.labels[prop] = 'int'
                        if prop in self.lens:
                            if len(str(attr)) > self.lens[prop][0]:
                                self.lens[prop][0] = len(str(attr))
                        else:
                            self.lens[prop] = [len(str(attr)), 0]
                    elif type(attr) is types.FloatType:
                        if self.labels[prop] == 'str':
                            self.labels[prop] = 'float'
                        a = str(attr)
                        bits = a.split('.')
                        if prop in self.lens:
                            for i in range(2):
                                if len(bits[i]) > self.lens[prop][i]:
                                    self.lens[prop][i] = len(bits[i])
                        else:
                            self.lens[prop] = [len(bits[0]), len(bits[1])]
        self.name = 'name'
        if self.fields is None:
            if self.name in self.labels:
                self.fields = [self.name]
            elif 'hour' in self.labels:
                self.name = 'hour'
                self.fields = [self.name]
            elif 'day' in self.labels:
                self.name = 'day'
                self.fields = [self.name]
            elif 'period' in self.labels:
                self.name = 'period'
                self.fields = [self.name]
            else:
                self.fields = []
            for key, value in iter(sorted(self.labels.iteritems())):
                if key != self.name:
                    self.fields.append(key)
        else:
            if 'name' not in self.fields:
                self.name = self.fields[0]
        self.entry = []
        if self.sumfields is not None:
            self.sums = {}
        iam = getattr(self.objects[0], '__module__')
        for obj in range(len(self.objects)):
            if not self.fossil:
                try:
                    attr = getattr(self.objects[obj], 'technology')
                    if attr[:6] == 'Fossil':
                        continue
                except:
                    pass
            if iam == 'Grid':
                attr = getattr(self.objects[obj], 'length')
                if attr < 0:
                    continue
            values = {}
            for key, value in self.labels.iteritems():
                try:
                    if key == '#':
                        attr = obj
                    elif key[0] == '%':
                        attr = ''
                    else:
                        attr = getattr(self.objects[obj], key)
                    if value != 'str':
                        fmat_str = '{: >' + str(self.lens[key][0] + self.lens[key][1] + 1) + ',.' + str(self.lens[key][1]) + 'f}'
                        values[key] = fmat_str.format(attr)
                    else:
                        values[key] = attr
                except:
                    pass
            if self.sumfields is not None or self.sumby is not None:
                sums = []
                for i in range(len(self.sumfields)):
                    if getattr(self.objects[obj], self.sumfields[i]) is None:
                        sums.append(0)
                    else:
                        sums.append(getattr(self.objects[obj], self.sumfields[i]))
                if self.sumby is not None:
                    sumby_key = getattr(self.objects[obj], self.sumby[0])
                    if sumby_key in self.sums:
                        csums = self.sums[sumby_key]
                        for i in range(len(self.sumfields)):
                            csums[i] = csums[i] + sums[i]
                        self.sums[sumby_key] = csums[:]
                    else:
                        self.sums[sumby_key] = sums[:]
                sumby_key = '~~'
                if sumby_key in self.sums:
                    tsums = self.sums[sumby_key]
                    for i in range(len(self.sumfields)):
                        try:
                            sums[i] = tsums[i] + sums[i]
                        except:
                            pass
                self.sums[sumby_key] = sums[:]
            self.entry.append(values)
        return len(self.objects)

    def order(self, col):
        key = self.fields[col]
        numbrs = True
        orderd = {}
        norderd = {}   # minus
        if 1 == 2:   # key == self.name:
            for rw in range(len(self.entry)):
                orderd[self.entry[rw][key]] = rw
        else:
            if key != '#':
                max_l = 0
                for rw in range(len(self.entry)):
                    if key in self.entry[rw]:
                        try:
                            max_l = max(max_l, len(self.entry[rw][key]))
                            try:
                                txt = str(self.entry[rw][key]).strip().replace(',', '')
                                nmbr = float(txt)
                            except:
                                numbrs = False
                        except:
                            pass
            if numbrs:
                fmat_str = '{:0>' + str(self.lens[key][0] + self.lens[key][1] + 1) + '.' + str(self.lens[key][1]) + 'f}'
            for rw in range(len(self.entry)):
                if key == '#':
                    orderd[str(rw).zfill(self.lens['#'][0]) + self.entry[rw][self.name]] = rw
                else:
                    if numbrs:
                        if key in self.entry[rw]:
                            txt = str(self.entry[rw][key]).strip().replace(',', '')
                            nmbr = float(txt)
                            if nmbr == 0:
                                orderd['<' + fmat_str.format(nmbr) + self.entry[rw][self.name]] = rw
                            elif nmbr < 0:
                                norderd['<' + fmat_str.format(-nmbr) + self.entry[rw][self.name]] = rw
                            else:
                                orderd['>' + fmat_str.format(nmbr) + self.entry[rw][self.name]] = rw
                        else:
                            orderd[' ' + self.entry[rw][self.name]] = rw
                    else:
                        try:
                            orderd[str(self.entry[rw][key]) + self.entry[rw][self.name]] = rw
                        except:
                            orderd[' ' + self.entry[rw][self.name]] = rw
        torder = []
        if col != self.sort_col:
            self.table.horizontalHeaderItem(self.sort_col).setIcon(QtGui.QIcon('blank.png'))
            self.sort_asc = False
        self.sort_col = col
        if self.sort_asc:   # swap order
            for key, value in iter(sorted(orderd.iteritems(), reverse=True)):
                torder.append(value)
            for key, value in iter(sorted(norderd.iteritems())):
                torder.append(value)
            self.sort_asc = False
            self.table.horizontalHeaderItem(col).setIcon(QtGui.QIcon('arrowd.png'))
        else:
            self.sort_asc = True
            for key, value in iter(sorted(norderd.iteritems(), reverse=True)):
                torder.append(value)
            for key, value in iter(sorted(orderd.iteritems())):
                torder.append(value)
            self.table.horizontalHeaderItem(col).setIcon(QtGui.QIcon('arrowu.png'))
        self.entry = [self.entry[i] for i in torder]
        for rw in range(len(self.entry)):
            for cl in range(self.table.columnCount()):
                self.table.setItem(rw, cl, QtGui.QTableWidgetItem(''))
            for key, value in sorted(self.entry[rw].items(), key=lambda i: self.fields.index(i[0])):
                cl = self.fields.index(key)
                if key == 'technology':
                    icon = self.icons.getIcon(value)
                    icon_item = QtGui.QTableWidgetItem(value)
                    icon_item.setIcon(QtGui.QIcon(icon))
                    self.table.setItem(rw, cl, QtGui.QTableWidgetItem(icon_item))
                elif key == 'coordinates':
                    if len(value) > 2:
                        trick = '(%s, %s)+%s+(%s, %s)' % (value[0][0], value[0][1], \
                                str(len(value) - 2), value[-1][0], value[-1][1])
                    else:
                        trick = '(%s, %s)(%s, %s)' % (value[0][0], value[0][1], \
                                value[-1][0], value[-1][1])
                    self.table.setItem(rw, cl, QtGui.QTableWidgetItem(trick))
                else:
                    if value is not None:
                        if isinstance(value, list):
                            fld = str(value[0])
                            for i in range(1, len(value)):
                                fld = fld + ',' + str(value[i])
                            self.table.setItem(rw, cl, QtGui.QTableWidgetItem(fld))
                        else:
                            self.table.setItem(rw, cl, QtGui.QTableWidgetItem(value))
                            if self.labels[key] != 'str':
                                self.table.item(rw, cl).setTextAlignment(130)   # x'82'
                if key == self.name or not self.edit_table:
                    self.table.item(rw, cl).setFlags(QtCore.Qt.ItemIsEnabled)

    def showit(self):
        self.show()

    def header_click(self, position):
        column = self.headers.logicalIndexAt(position)
        self.order(column)

    def item_selected(self, row, col):
        for thing in self.objects:
            try:
                attr = getattr(thing, self.name)
                if attr == self.entry[row][self.name]:
                    dialog = displayobject.AnObject(QtGui.QDialog(), thing)
                    dialog.exec_()
                    break
            except:
                pass

    def item_changed(self, row, col):
        if self.recur:
            return
        self.message.setText(' ')
        if self.labels[self.fields[col]] == 'int' or self.labels[self.fields[col]] == 'float':
            self.recur = True
            tst = str(self.table.item(row, col).text().replace(',', ''))
            mult = 1
            if tst[-1].upper() == 'K':
                mult = 1 * pow(10, 3)
                tst = tst[:-1]
            elif tst[-1].upper() == 'M':
                mult = 1 * pow(10, 6)
                tst = tst[:-1]
            if self.labels[self.fields[col]] == 'int':
                try:
                    tst = int(tst) * mult
                    fmat_str = '{: >' + str(self.lens[self.fields[col]][0] + self.lens[self.fields[col]][1] + 1) + ',.' \
                               + str(self.lens[self.fields[col]][1]) + 'f}'
                    self.table.setItem(row, col, QtGui.QTableWidgetItem(fmat_str.format(tst)))
                    self.table.item(row, col).setTextAlignment(130)  # x'82'
                    self.recur = False
                    return
                except:
                    self.message.setText('Error with ' + self.fields[col].title() + ' field - ' + tst)
                    self.recur = False
            else:
                try:
                    tst = float(tst) * mult
                    fmat_str = '{: >' + str(self.lens[self.fields[col]][0] + self.lens[self.fields[col]][1] + 1) + ',.' \
                               + str(self.lens[self.fields[col]][1]) + 'f}'
                    self.table.setItem(row, col, QtGui.QTableWidgetItem(fmat_str.format(tst)))
                    self.table.item(row, col).setTextAlignment(130)  # x'82'
                    self.recur = False
                    return
                except:
                    self.message.setText('Error with ' + self.fields[col].title() + ' field - ' + tst)
                    self.recur = False
        msg_font = self.message.font()
        msg_font.setBold(True)
        self.message.setFont(msg_font)
        msg_palette = QtGui.QPalette()
        msg_palette.setColor(QtGui.QPalette.Foreground, QtCore.Qt.red)
        self.message.setPalette(msg_palette)

    def quit(self):
        self.close()

    def saveit(self):
        if self.title is None:
            iam = getattr(self.objects[0], '__module__')
        else:
            iam = self.title
        data_file = '%s_Table_%s.xls' % (iam, \
                    str(QtCore.QDateTime.toString(QtCore.QDateTime.currentDateTime(), 'yyyy-MM-dd_hhmm')))
        data_file = QtGui.QFileDialog.getSaveFileName(None, 'Save ' + iam + ' Table',
                    self.save_folder + data_file, 'Excel Files (*.xls*);;CSV Files (*.csv)')
        if data_file == '':
            return
        if data_file[-4:] == '.csv' or data_file[-4:] == '.xls' or data_file[-5:] == '.xlsx':
            pass
        else:
            data_file += '.xls'
        if os.path.exists(data_file):
            if os.path.exists(data_file + '~'):
                os.remove(data_file + '~')
            os.rename(data_file, data_file + '~')
        if data_file[-4:] == '.csv':
            tf = open(data_file, 'w')
            hdr_types = []
            line = ''
            for cl in range(self.table.columnCount()):
                if cl > 0:
                    line += ','
                txt = str(self.table.horizontalHeaderItem(cl).text())
                if ',' in txt:
                    line += '"' + txt + '"'
                else:
                    line += txt
                hdr_types.append(self.labels[txt.lower()])
            tf.write(line + '\n')
            for rw in range(self.table.rowCount()):
                line = ''
                for cl in range(self.table.columnCount()):
                    if cl > 0:
                        line += ','
                    if self.table.item(rw, cl) is not None:
                        txt = str(self.table.item(rw, cl).text())
                        if hdr_types[cl] == 'int':
                            try:
                                txt = str(self.table.item(rw, cl).text()).strip()
                            except:
                                pass
                        elif hdr_types[cl] == 'float':
                            try:
                                txt = str(self.table.item(rw, cl).text()).strip()
                                txt = val1.replace(',', '')
                            except:
                                pass
                        if ',' in txt:
                            line += '"' + txt + '"'
                        else:
                            line += txt
                tf.write(line + '\n')
            tf.close()
        else:
            wb = xlwt.Workbook()
            ws = wb.add_sheet(iam)
            hdr_types = []
            dec_fmts = []
            xl_lens = []
            for cl in range(self.table.columnCount()):
                hdr = str(self.table.horizontalHeaderItem(cl).text())
                if hdr[0] != '%':
                    ws.write(0, cl, hdr)
                txt = self.hdrs[hdr]
                hdr_types.append(self.labels[txt.lower()])
                style = xlwt.XFStyle()
                try:
                    if self.lens[txt.lower()][1] > 0:
                        style.num_format_str = '#,##0.' + '0' * self.lens[txt.lower()][1]
                    elif self.labels[txt.lower()] == 'int':
                        style.num_format_str = '#,##0'
                except:
                    pass
                dec_fmts.append(style)
                xl_lens.append(len(hdr))
            for rw in range(self.table.rowCount()):
                for cl in range(self.table.columnCount()):
                    if self.table.item(rw, cl) is not None:
                        valu = str(self.table.item(rw, cl).text())
                        if hdr_types[cl] == 'int':
                            try:
                                val1 = str(self.table.item(rw, cl).text()).strip()
                                val1 = val1.replace(',', '')
                                valu = int(val1)
                            except:
                                pass
                        elif hdr_types[cl] == 'float':
                            try:
                                val1 = str(self.table.item(rw, cl).text()).strip()
                                val1 = val1.replace(',', '')
                                valu = float(val1)
                            except:
                                pass
                        xl_lens[cl] = max(xl_lens[cl], len(str(valu)))
                        ws.write(rw + 1, cl, valu, dec_fmts[cl])
            for cl in range(self.table.columnCount()):
                if xl_lens[cl] * 275 > ws.col(cl).width:
                    ws.col(cl).width = xl_lens[cl] * 275
            ws.set_panes_frozen(True)   # frozen headings instead of split panes
            ws.set_horz_split_pos(1)   # in general, freeze after last heading row
            ws.set_remove_splits(True)   # if user does unfreeze, don't leave a split there
            wb.save(data_file)
        self.close()

    def replacetbl(self):
        self.replaced = {}
        for rw in range(self.table.rowCount()):
            key = str(self.table.item(rw, 0).text())
            values = []
            for cl in range(1, self.table.columnCount()):
                if self.table.item(rw, cl) is not None:
                    if self.labels[self.fields[cl]] == 'int' or self.labels[self.fields[cl]] == 'float':
                        tst = str(self.table.item(rw, cl).text()).strip().replace(',', '')
                        if tst == '':
                            valu = tst
                        else:
                            mult = 1
                            if tst[-1].upper() == 'K':
                                mult = 1 * pow(10, 3)
                                tst = tst[:-1]
                            elif tst[-1].upper() == 'M':
                                mult = 1 * pow(10, 6)
                                tst = tst[:-1]
                            if self.labels[self.fields[cl]] == 'int':
                                try:
                                    valu = int(tst) * mult
                                    if valu == 0:
                                        valu = ''
                                    elif valu > 99 and valu < 100000:
                                        if len(str(valu / pow(10, 3)).split('.')[1]) < 3:
                                            valu = str(valu / pow(10, 3)) + 'K'
                                        else:
                                            value = str(valu)
                                    elif valu >= 100000:
                                        if len(str(valu / pow(10, 6)).split('.')[1]) < 3:
                                            valu = str(valu / pow(10, 6)) + 'M'
                                        else:
                                            valu = str(valu)
                                    else:
                                        valu = str(valu)
                                except:
                                    self.message.setText('Error with ' + self.fields[cl].title() + ' field - ' + tst)
                                    self.replaced = None
                                    return
                            else:
                                try:
                                    valu = float(tst) * mult
                                    if valu == 0:
                                        valu = ''
                                    elif len(str(valu).split('.')[1]) > 1:
                                        valu = str(valu)
                                    elif valu > 99 and valu < 1000000:
                                        if len(str(valu / pow(10, 3)).split('.')[1]) < 3:
                                            valu = str(valu / pow(10, 3)) + 'K'
                                            valu = valu.replace('.0K', 'K')
                                        else:
                                            value = str(valu)
                                    elif valu >= 1000000:
                                        if len(str(valu / pow(10, 6)).split('.')[1]) < 3:
                                            valu = str(valu / pow(10, 6)) + 'M'
                                            valu = valu.replace('.0M', 'M')
                                        else:
                                            valu = str(valu)
                                    else:
                                        valu = str(valu)
                                except:
                                    self.message.setText('Error with ' + self.fields[cl].title() + ' field - ' + tst)
                                    self.replace = None
                                    return
                    else:
                        valu = str(self.table.item(rw, cl).text())
                else:
                    valu = ''
                if valu != '':
                    values.append(self.fields[cl] + '=' + valu)
            self.replaced[key] = values
        self.close()

    def getValues(self):
        if self.edit_table:
            return self.replaced
        return None
