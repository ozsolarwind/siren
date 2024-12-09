#!/usr/bin/python3
#
#  Copyright (C) 2015-2024 Sustainable Energy Now Inc., Angus King
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

import openpyxl as oxl
import os
import xlwt
from PyQt5 import QtCore, QtWidgets
from PyQt5 import QtGui, QtWidgets

import displayobject
from sirenicons import Icons
from senutils import ssCol, techClean


class FakeObject:
    def __init__(self, fake_object, fields):
        f = -1
        if not isinstance(fake_object, list) and len(fields) > 1:
            f += 1
            setattr(self, fields[f], fake_object)
            for f in range(1, len(fields)):
                setattr(self, fields[f], '')
            return
        for i in range(len(fake_object)):
            if isinstance(fake_object[i], list):
                for j in range(len(fake_object[i])):
                    f += 1
                    setattr(self, fields[f], fake_object[i][j])
            else:
                f += 1
                setattr(self, fields[f], fake_object[i])


class Table(QtWidgets.QDialog):
    def __init__(self, objects, parent=None, fields=None, fossil=True, sumby=None, sumfields=None, units='', title=None,
                 save_folder='', edit=False, sortby=None, decpts=None, totfields=None, abbr=True, txt_align=None,
                 reverse=False, txt_ok=None, span=None, year=''):
        super(Table, self).__init__(parent)
        self.oclass = None
        if len(objects) == 0:
            buttonLayout = QtWidgets.QVBoxLayout()
            buttonLayout.addWidget(QtWidgets.QLabel('Nothing to display.'))
            self.quitButton = QtWidgets.QPushButton(self.tr('&Quit'))
            buttonLayout.addWidget(self.quitButton)
            self.quitButton.clicked.connect(self.quit)
            QtWidgets.QShortcut(QtGui.QKeySequence('q'), self, self.quit)
            self.setLayout(buttonLayout)
            return
        elif isinstance(objects, list) and isinstance(objects[0], list):
            fakes = []
            for row in objects:
                fakes.append(FakeObject(row, fields))
            self.objects = fakes
        elif isinstance(objects, dict):
            fakes = []
            if fields is None: # assume we have some class objects
                self.oclass = type(objects[list(objects.keys())[0]])
                fields = []
                if hasattr(objects[list(objects.keys())[0]], 'name'):
                    fields.append('name')
                for prop in dir(objects[list(objects.keys())[0]]):
                    if prop[:2] != '__' and prop[-2:] != '__':
                        if prop != 'name':
                            fields.append(prop)
                for key, value in objects.items():
                     values = []
                     for field in fields:
                         if isinstance(getattr(value, field), list):
                             txt = ''
                             fld = getattr(value, field)
                             for fl in fld:
                                 txt += str(fl) + ' '
                             txt = txt[:-2]
                             values.append(txt)
                         else:
                             values.append(getattr(value, field))
                     fakes.append(FakeObject(values, fields))
            else:
                for key, value in objects.items():
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
        self.totfields = totfields
        self.units = units
        self.title = title
        self.edit_table = False
        self.edit_delete = False
        if edit:
            self.edit_table = edit
            try:
                if edit.lower() == 'delete':
                    self.edit_delete = True
            except:
                pass
        self.decpts = decpts
        self.txt_align = txt_align
        self.txt_ok = txt_ok
        self.recur = False
        self.replaced = None
        self.savedfile = None
        self.span = span
        if year != '':
            self.year = year + '_'
        else:
            self.year = year
        self.abbr = abbr
        if self.edit_table:
            if self.edit_delete:
                self.title_word = ['List', 'Export']
            else:
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
            try:
                self.setWindowTitle('SIREN - ' + self.title_word[0] + ' ' + getattr(objects[0], '__module__') + 's')
            except:
                self.setWindowTitle('SIREN - ' + self.title_word[0] + 'items')
        else:
            self.setWindowTitle('SIREN - ' + self.title_word[0] + ' ' + self.title)
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        msg = '(Right click column header to sort)'
        if self.edit_table and (self.fields[0] == 'name' or self.edit_delete):
            msg = msg[:-1] + '; right click row number to delete)'
        try:
            if getattr(objects[0], '__module__') == 'Station':
                msg = '(Left or right click row to display, ' + msg[1:]
        except:
            pass
        buttonLayout = QtWidgets.QHBoxLayout()
        self.message = QtWidgets.QLabel(msg)
        self.quitButton = QtWidgets.QPushButton(self.tr('&Quit'))
        buttonLayout.addWidget(self.quitButton)
        self.quitButton.clicked.connect(self.quit)
        if self.edit_table:
            if not self.edit_delete:
                if isinstance(objects, dict):
                    if fields[0] == 'property' or fields[0] == 'name':
                        self.addButton = QtWidgets.QPushButton(self.tr('Add'))
                        buttonLayout.addWidget(self.addButton)
                        self.addButton.clicked.connect(self.addtotbl)
            self.replaceButton = QtWidgets.QPushButton(self.tr('Save'))
            buttonLayout.addWidget(self.replaceButton)
            self.replaceButton.clicked.connect(self.replacetbl)
        self.saveButton = QtWidgets.QPushButton(self.tr(self.title_word[1]))
        buttonLayout.addWidget(self.saveButton)
        self.saveButton.clicked.connect(self.saveit)
        buttons = QtWidgets.QFrame()
        buttons.setLayout(buttonLayout)
        layout = QtWidgets.QGridLayout()
        self.table = QtWidgets.QTableWidget()
        self.populate()
        if self.sumfields is None:
            self.table.setRowCount(len(self.entry))
        else:
            self.table.setRowCount(len(self.entry) + len(self.sums))
        self.table.setColumnCount(len(self.labels))
        if self.fields is None:
            labels = sorted([self.nice(x) for x in list(self.labels.keys())])
        else:
            for f in range(len(self.fields) -1, -1, -1):
                if self.fields[f] not in list(self.labels.keys()):
                    del self.fields[f] # delete any None fields
            labels = [self.nice(x) for x in self.fields]
        self.table.setHorizontalHeaderLabels(labels)
        for cl in range(self.table.columnCount()):
            self.table.horizontalHeaderItem(cl).setIcon(QtGui.QIcon('blank.png'))
        self.headers = self.table.horizontalHeader()
        self.headers.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.headers.customContextMenuRequested.connect(self.header_click)
        self.headers.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.table.verticalHeader().setVisible(False)
      #   self.table.setEditTriggers(QtGui.QAbstractItemView.NoEditTriggers)
        if self.edit_table:
            self.table.setEditTriggers(QtWidgets.QAbstractItemView.CurrentChanged)
            if self.fields[0] == 'name' or self.edit_delete:
                self.rows = self.table.verticalHeader()
                self.rows.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
                self.rows.customContextMenuRequested.connect(self.row_click)
                self.rows.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
                self.table.verticalHeader().setVisible(True)
        else:
            self.table.setEditTriggers(QtWidgets.QAbstractItemView.SelectedClicked)
            self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        layout.addWidget(self.table, 0, 0)
        layout.addWidget(self.message, 1, 0)
        layout.addWidget(buttons, 2, 0)
        self.sort_col = 0
        self.sort_asc = False
        if sortby is None:
            self.order(0)
        elif sortby == '':
            self.order(-1)
        else:
            self.order(self.fields.index(sortby))
            if reverse:
                self.order(self.fields.index(sortby))
        if self.sumfields is not None:
            for i in range(len(self.sumfields) -1, -1, -1): # make sure sumfield has a field
                try:
                    self.fields.index(self.sumfields[i])
                except:
                    del self.sumfields[i]
            fmat_str = []
            clv = []
            clp = []
            totl = []
            for i in range(len(self.sumfields)):
                clp.append(-1)
            for f in range(len(self.sumfields)):
                fmat_str.append('{:,.' + str(self.lens[self.sumfields[f]][1]) + 'f}')
                clv.append(self.fields.index(self.sumfields[f]))
            if self.totfields is not None:
                for i in range(len(self.totfields)):
               #     clp.append(-1)
                    fmat_str.append('{: >' + str(self.lens[self.totfields[i][0]][0] + \
                                    self.lens[self.totfields[i][0]][1] + 1) + ',.' + \
                                    str(self.lens[self.totfields[i][0]][1]) + 'f}')
                    clv.append(self.fields.index(self.totfields[i][0]))
            if self.sumby is not None:
                clk = self.fields.index(self.sumby[0])
                for i in range(len(self.sumfields)):
                    clp[i] = clv[i] + 1
                    totl.append(self.sums['~~'][i])
            rw = len(self.entry) - 1
            for key, value in iter(sorted(self.sums.items())):
                if key == '~~':
                    tkey = 'Total'
                else:
                    tkey = key
                rw += 1
                if self.sumby is not None:
                    if self.sumby[0] == 'technology' and tkey != 'Total':
                        icon = self.icons.getIcon(tkey)
                        icon_item = QtWidgets.QTableWidgetItem(str(tkey))
                        icon_item.setIcon(QtGui.QIcon(icon))
                        self.table.setItem(rw, clk, QtWidgets.QTableWidgetItem(icon_item))
                    else:
                        self.table.setItem(rw, clk, QtWidgets.QTableWidgetItem(str(tkey)))
                        if not isinstance(tkey, str):
                            self.table.item(rw, clk).setTextAlignment(130)  # x'82'
                for f in range(len(self.sumfields)):
                    self.table.setItem(rw, clv[f], QtWidgets.QTableWidgetItem(fmat_str[f].format(value[f])))
                    self.table.item(rw, clv[f]).setTextAlignment(130)  # x'82'
                    if clp[f] > 0 and totl[f] > 0:
                        self.table.setItem(rw, clp[f], QtWidgets.QTableWidgetItem('{:.1%}'.format(float(value[f]) /
                          float(totl[f])) + ' '))
                        self.table.item(rw, clp[f]).setTextAlignment(130)  # x'82'
            if self.totfields is not None:
                for f in range(len(self.totfields)):
                    self.table.setItem(rw, clv[len(self.sumfields) + f],
                        QtWidgets.QTableWidgetItem(fmat_str[len(self.sumfields) + f].format(self.totfields[f][1])))
                    self.table.item(rw, clv[len(self.sumfields) + f]).setTextAlignment(130)  # x'82'
        if self.span is None:
            self.table.resizeColumnsToContents()
        width = 0
        for cl in range(self.table.columnCount()):
            width += self.table.columnWidth(cl)
        width += 50
        height = self.table.rowHeight(1) * (self.table.rowCount() + 4)
        screen = QtWidgets.QDesktopWidget().availableGeometry()
        if height > (screen.height() - 70):
            height = screen.height() - 70
        self.setLayout(layout)
        size = QtCore.QSize(QtCore.QSize(int(width), int(height)))
        self.resize(size)
        self.updated = QtCore.pyqtSignal(QtWidgets.QLabel)   # ??
        QtWidgets.QShortcut(QtGui.QKeySequence('q'), self, self.quit)
        if self.edit_table and not self.edit_delete:
            self.table.cellChanged.connect(self.item_changed)
        else:
            self.table.cellClicked.connect(self.item_selected)

    def nice(self, string):
        if string == '':
            strout = string
        else:
            strout = techClean(string, full=True)
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
        try:
            self.hdrs
        except:
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
                        if isinstance(attr, int):
                            self.labels[prop] = 'int'
                        elif isinstance(attr, float):
                            self.labels[prop] = 'float'
                        else:
                            try:
                                at = float(attr.strip('%'))
                                self.labels[prop] = 'float'
                            except:
                                self.labels[prop] = 'str'
                    if isinstance(attr, int):
                        if self.labels[prop] == 'str':
                            self.labels[prop] = 'int'
                        if prop in self.lens:
                            if len(str(attr)) > self.lens[prop][0]:
                                self.lens[prop][0] = len(str(attr))
                        else:
                            self.lens[prop] = [len(str(attr)), 0]
                    elif isinstance(attr, float):
                        if self.labels[prop] == 'str' or self.labels[prop] == 'int':
                            self.labels[prop] = 'float'
                        a = str(attr)
                        bits = a.split('.')
                        if len(bits) == 1: # maybe other float format
                            bits = a.split('e')
                        if self.decpts is None:
                            if prop in self.lens:
                                for i in range(2):
                                    if len(bits[i]) > self.lens[prop][i]:
                                        self.lens[prop][i] = len(bits[i])
                            else:
                                self.lens[prop] = [len(bits[0]), len(bits[1])]
                        else:
                            pts = self.decpts[self.fields.index(prop)]
                            if prop in self.lens:
                                if len(bits[0]) > self.lens[prop][0]:
                                    self.lens[prop][0] = len(bits[0])
                                if len(bits[1]) > self.lens[prop][1]:
                                    if len(bits[1]) > pts:
                                        self.lens[prop][1] = pts
                                    else:
                                        self.lens[prop][1] = len(bits[1])
                            else:
                                if len(bits[1]) > pts or self.edit_table:
                                    self.lens[prop] = [len(bits[0]), pts]
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
            for key, value in iter(sorted(self.labels.items())):
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
            for key, value in self.labels.items():
                attr = ''
                try:
                    if key == '#':
                        attr = obj
                    elif key[0] == '%':
                        attr = ''
                    else:
                        attr = getattr(self.objects[obj], key)
                    if value != 'str':
                        if isinstance(attr, str):
                            if attr == '':
                                continue
                            values[key] = attr
                        else:
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
        if col < 0:   # key == self.name:
            torder = []
            for rw in range(len(self.entry)):
                torder.append(rw)
        else:
            numbrs = True
            orderd = {}
            norderd = {}   # minus
            key = self.fields[col]
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
            # a quirk here is that reverse order will have the name field in reverse order for equal
            # sort column values. A might look at fixing it someday
            if self.sort_asc:   # swap order
                for key, value in iter(sorted(iter(orderd.items()), reverse=True)):
                    torder.append(value)
                for key, value in iter(sorted(norderd.items())):
                    torder.append(value)
                self.table.horizontalHeaderItem(col).setIcon(QtGui.QIcon('arrowd.png'))
                self.sort_asc = False
            else:
                for key, value in iter(sorted(iter(norderd.items()), reverse=True)):
                    torder.append(value)
                for key, value in iter(sorted(orderd.items())):
                    torder.append(value)
                self.table.horizontalHeaderItem(col).setIcon(QtGui.QIcon('arrowu.png'))
                self.sort_asc = True
        self.entry = [self.entry[i] for i in torder]
        in_span = False
        for rw in range(len(self.entry)):
            if self.span is not None and not in_span:
                self.table.resizeColumnsToContents()
            for key, value in sorted(list(self.entry[rw].items()), key=lambda i: self.fields.index(i[0])):
                cl = self.fields.index(key)
                if cl == 0:
                    if self.span is not None and value == self.span:
                        in_span = True
                if key == 'technology':
                    icon = self.icons.getIcon(value)
                    icon_item = QtWidgets.QTableWidgetItem(value)
                    icon_item.setIcon(QtGui.QIcon(icon))
                    self.table.setItem(rw, cl, QtWidgets.QTableWidgetItem(icon_item))
                elif key == 'coordinates':
                    if len(value) > 2:
                        trick = '(%s, %s)+%s+(%s, %s)' % (value[0][0], value[0][1],
                                str(len(value) - 2), value[-1][0], value[-1][1])
                    else:
                        trick = '(%s, %s)(%s, %s)' % (value[0][0], value[0][1],
                                value[-1][0], value[-1][1])
                    self.table.setItem(rw, cl, QtWidgets.QTableWidgetItem(trick))
                else:
                    if value is not None:
                        if isinstance(value, list):
                            fld = str(value[0])
                            for i in range(1, len(value)):
                                fld = fld + ',' + str(value[i])
                            self.table.setItem(rw, cl, QtWidgets.QTableWidgetItem(fld))
                            if self.txt_align is not None:
                                if self.txt_align == 'R':
                                    self.table.item(rw, cl).setTextAlignment(130)  # x'82'
                        else:
                            if cl == 1 and in_span:
                                self.table.setSpan(rw, cl, 1, len(self.fields) - 1)
                                self.table.setItem(rw, cl, QtWidgets.QTableWidgetItem(value))
                                continue
                            try:
                                self.table.setItem(rw, cl, QtWidgets.QTableWidgetItem(value))
                            except: # allow for other things e.g. Widgets
                                self.table.setCellWidget(rw, cl, value)
                            if self.labels[key] != 'str' or \
                               (self.txt_align is not None and self.txt_align == 'R'):
                                self.table.item(rw, cl).setTextAlignment(130)   # x'82'
                    else:
                        self.table.setItem(rw, cl, QtWidgets.QTableWidgetItem(''))
         #       if key == self.name or not self.edit_table:
                if not self.edit_table:
                    self.table.item(rw, cl).setFlags(QtCore.Qt.ItemIsEnabled)

    def showit(self):
        self.show()

    def header_click(self, position):
        column = self.headers.logicalIndexAt(position)
        self.order(column)

    def row_click(self, position):
        row = self.rows.logicalIndexAt(position)
        msgbox = QtWidgets.QMessageBox()
        msgbox.setWindowTitle('SIREN - Delete item')
        msgbox.setText("Press Yes to delete '" + self.table.item(row, 0).text() + "'")
        msgbox.setIcon(QtWidgets.QMessageBox.Question)
        msgbox.setStandardButtons(QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
        reply = msgbox.exec_()
        if reply == QtWidgets.QMessageBox.Yes:
            for i in range(len(self.entry)):
                if self.edit_delete:
                    if self.entry[i][self.fields[0]] == self.table.item(row, 0).text():
                        del self.entry[i]
                        break
                elif self.entry[i]['name'] == self.table.item(row, 0).text():
                    del self.entry[i]
                    break
            for i in range(len(self.objects)):
                if self.edit_delete:
                    value = getattr(self.objects[i], self.fields[0])
                    if value == self.table.item(row, 0).text():
                        del self.objects[i]
                        break
                elif self.objects[i].name == self.table.item(row, 0).text():
                    del self.objects[i]
                    break
            self.table.removeRow(row)

    def addtotbl(self):
        addproperty = {}
        for field in self.fields:
            addproperty[field] = ''
        if self.fields[0] == 'name':
            textedit = False
        else:
            textedit = True
        dialog = displayobject.AnObject(QtWidgets.QDialog(), addproperty, readonly=False,
                 textedit=textedit, title='Add ' + self.fields[0].title())
        dialog.exec_()
        if dialog.getValues()[self.fields[0]] != '':
            self.entry.append(addproperty)
            if self.fields[0] == 'property':
                self.objects.append(FakeObject([dialog.getValues()[self.fields[0]],
                                    dialog.getValues()[self.fields[1]]], self.fields))
            else:
                fakes = []
                for field in self.fields:
                    if self.labels[field] == 'int':
                        try:
                            fakes.append(int(dialog.getValues()[field]))
                        except:
                            fakes.append(0)
                    elif self.labels[field] == 'float':
                        try:
                            fakes.append(float(dialog.getValues()[field]))
                        except:
                            fakes.append(0.)
                    else:
                        fakes.append(dialog.getValues()[field])
                self.objects.append(FakeObject(fakes, self.fields))
            self.populate()
            self.table.setRowCount(self.table.rowCount() + 1)
            self.sort_col = 1
            self.recur = True
            self.order(0)
            self.recur = False
        del dialog

    def item_selected(self, row, col):
        for thing in self.objects:
            try:
                attr = getattr(thing, self.name)
                if attr == self.entry[row][self.name]:
                    dialog = displayobject.AnObject(QtWidgets.QDialog(), thing)
                    dialog.exec_()
                    break
            except:
                pass

    def item_changed(self, row, col):
        if self.recur:
            return
        self.entry[row][self.fields[col]] = self.table.item(row, col).text()
        self.message.setText(' ')
        if self.labels[self.fields[col]] == 'int' or self.labels[self.fields[col]] == 'float':
            self.recur = True
            tst = self.table.item(row, col).text().replace(',', '')
            if len(tst) < 1:
                return
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
                    fmat_str = '{: >' + str(self.lens[self.fields[col]][0] \
                               + self.lens[self.fields[col]][1] + 1) + ',.' \
                               + str(self.lens[self.fields[col]][1]) + 'f}'
                    self.table.setItem(row, col, QtWidgets.QTableWidgetItem(fmat_str.format(tst)))
                    self.table.item(row, col).setTextAlignment(130)  # x'82'
                    self.recur = False
                    return
                except:
                    if self.txt_ok is not None and self.table.item(row, 0).text() in self.txt_ok:
                        self.recur = False
                        return
                    self.message.setText('Error with ' + self.fields[col].title() + ' field - ' + tst)
                    self.recur = False
            else:
                try:
                    tst = float(tst) * mult
                    fmat_str = '{: >' + str(self.lens[self.fields[col]][0] \
                               + self.lens[self.fields[col]][1] + 1) + ',.' \
                               + str(self.lens[self.fields[col]][1]) + 'f}'
                    self.table.setItem(row, col, QtWidgets.QTableWidgetItem(fmat_str.format(tst)))
                    self.table.item(row, col).setTextAlignment(130)  # x'82'
                    self.recur = False
                    return
                except:
                    if self.txt_ok is not None and self.table.item(row, 0).text() in self.txt_ok:
                        self.recur = False
                        return
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
        data_file = '%s_Table_%s%s.xlsx' % (iam, self.year,
                    QtCore.QDateTime.toString(QtCore.QDateTime.currentDateTime(), 'yyyy-MM-dd_hhmm'))
        data_file = QtWidgets.QFileDialog.getSaveFileName(None, 'Save ' + iam + ' Table',
                    self.save_folder + data_file, 'Excel Files (*.xls*);;CSV Files (*.csv)')[0]
        if data_file == '':
            return
        if data_file[-4:] == '.csv' or data_file[-4:] == '.xls' or data_file[-5:] == '.xlsx':
            pass
        else:
            data_file += '.xlsx'
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
                hdr = self.table.horizontalHeaderItem(cl).text()
                if hdr[0] != '%':
                    txt = hdr
                    if ',' in txt:
                        line += '"' + txt + '"'
                    else:
                        line += txt
                txt = self.hdrs[hdr]
                try:
                    hdr_types.append(self.labels[txt.lower()])
                except:
                    hdr_types.append(self.labels[txt])
            tf.write(line + '\n')
            for rw in range(self.table.rowCount()):
                line = ''
                for cl in range(self.table.columnCount()):
                    if cl > 0:
                        line += ','
                    if self.table.item(rw, cl) is not None:
                        txt = self.table.item(rw, cl).text()
                        if hdr_types[cl] == 'int':
                            try:
                                txt = self.table.item(rw, cl).text().strip()
                            except:
                                pass
                        elif hdr_types[cl] == 'float':
                            try:
                                txt = self.table.item(rw, cl).text().strip()
                                txt = txt.replace(',', '')
                            except:
                                pass
                        if ',' in txt:
                            line += '"' + txt + '"'
                        else:
                            line += txt
                tf.write(line + '\n')
            tf.close()
        elif data_file[-4:] == '.xls':
            wb = xlwt.Workbook()
            for ch in ['\\' , '/' , '*' , '?' , ':' , '[' , ']']:
                if ch in iam:
                    iam = iam.replace(ch, '_')
            if len(iam) > 31:
                iam = iam[:31]
            ws = wb.add_sheet(iam)
            hdr_types = []
            dec_fmts = []
            xl_lens = []
            hdr_rows = 0
            hdr_style = xlwt.XFStyle()
            hdr_style.alignment.wrap = 1
            for cl in range(self.table.columnCount()):
                hdr = self.table.horizontalHeaderItem(cl).text()
                if hdr[0] != '%':
                    ws.write(0, cl, hdr, hdr_style)
                txt = self.hdrs[hdr]
                try:
                    hdr_types.append(self.labels[txt.lower()])
                    txt = txt.lower()
                except:
                    try:
                        hdr_types.append(self.labels[txt])
                    except:
                        hdr_types.append('str')
                style = xlwt.XFStyle()
                try:
                    if self.lens[txt][1] > 0:
                        style.num_format_str = '#,##0.' + '0' * self.lens[txt][1]
                    elif self.labels[txt] == 'int' or self.labels[txt] == 'float':
                        style.num_format_str = '#,##0'
                except:
                    pass
                dec_fmts.append(style)
                bits = hdr.split('\n')
                hdr_rows = max(hdr_rows, len(bits))
                hl = 0
                for bit in bits:
                    hl = max(hl, len(bit) + 1)
                xl_lens.append(hl)
            if hdr_rows > 1:
                ws.row(0).height = 250 * hdr_rows
            in_span = False
            for rw in range(self.table.rowCount()):
                for cl in range(self.table.columnCount()):
                    if self.table.item(rw, cl) is not None:
                        valu = self.table.item(rw, cl).text().strip()
                        if len(valu) < 1:
                            continue
                        if self.span is not None and valu == self.span:
                            in_span = True
                        style = dec_fmts[cl]
                        if valu[-1] == '%':
                            is_pct = True
                            i = valu.rfind('.')
                            if i >= 0:
                                dec_pts = (len(valu) - i - 2)
                                style = xlwt.XFStyle()
                                try:
                                    style.num_format_str = '#,##0.' + '0' * dec_pts + '%'
                                except:
                                    pass
                            else:
                                 dec_pts = 0
                                 style.num_format_str = '#,##0%'
                        else:
                            is_pct = False
                        if hdr_types[cl] == 'int':
                            try:
                                val1 = valu
                                if is_pct:
                                    val1 = val1.strip('%')
                                val1 = val1.replace(',', '')
                                if is_pct:
                                    valu = round(int(val1) / 100., dec_pts + 2)
                                else:
                                    valu = int(val1)
                            except:
                                pass
                        elif hdr_types[cl] == 'float':
                            try:
                                val1 = valu
                                if is_pct:
                                    val1 = val1.strip('%')
                                val1 = val1.replace(',', '')
                                if is_pct:
                                    valu = round(float(val1) / 100., dec_pts + 2)
                                else:
                                    valu = float(val1)
                            except:
                                pass
                        else:
                            if is_pct:
                                try:
                                    val1 = valu.strip('%')
                                    val1 = val1.replace(',', '')
                                    valu = round(float(val1) / 100., dec_pts + 2)
                                except:
                                    pass
                        if not in_span:
                            xl_lens[cl] = max(xl_lens[cl], len(str(valu)))
                        ws.write(rw + 1, cl, valu, style)
            for cl in range(self.table.columnCount()):
                if xl_lens[cl] * 275 > ws.col(cl).width:
                    ws.col(cl).width = xl_lens[cl] * 275
            ws.set_panes_frozen(True)   # frozen headings instead of split panes
            ws.set_horz_split_pos(1)   # in general, freeze after last heading row
            ws.set_remove_splits(True)   # if user does unfreeze, don't leave a split there
            wb.save(data_file)
        else: # .xlsx
            wb = oxl.Workbook()
            ws = wb.active
            for ch in ['\\' , '/' , '*' , '?' , ':' , '[' , ']']:
                if ch in iam:
                    iam = iam.replace(ch, '_')
            if len(iam) > 31:
                iam = iam[:31]
            ws.title = iam
            normal = oxl.styles.Font(name='Arial', size='10')
        #    bold = oxl.styles.Font(name='Arial', bold=True)
            hdr_types = []
            dec_fmts = []
            xl_lens = []
            hdr_rows = 0
        #    hdr_style = xlwt.XFStyle()
        #    hdr_style.alignment.wrap = 1
            for cl in range(self.table.columnCount()):
                hdr = self.table.horizontalHeaderItem(cl).text()
                if hdr[0] != '%':
                    ws.cell(row=1, column=cl + 1).value = hdr
                    ws.cell(row=1, column=cl + 1).font = normal
                    ws.cell(row=1, column=cl + 1).alignment = oxl.styles.Alignment(wrap_text=True,
                            vertical='bottom', horizontal='center')
                txt = self.hdrs[hdr]
                try:
                    hdr_types.append(self.labels[txt.lower()])
                    txt = txt.lower()
                except:
                    try:
                        hdr_types.append(self.labels[txt])
                    except:
                        hdr_types.append('str')
                style = ''
                try:
                    if self.lens[txt][1] > 0:
                        style = '#,##0.' + '0' * self.lens[txt][1]
                    elif self.labels[txt] == 'int' or self.labels[txt] == 'float':
                        style = '#,##0'
                except:
                    pass
                dec_fmts.append(style)
                bits = hdr.split('\n')
                hdr_rows = max(hdr_rows, len(bits))
                hl = 0
                for bit in bits:
                    hl = max(hl, len(bit) + 1)
                xl_lens.append(hl)
            if hdr_rows > 1:
                ws.row_dimensions[1].height = 12 * hdr_rows
            in_span = False
            for rw in range(self.table.rowCount()):
                for cl in range(self.table.columnCount()):
                    if self.table.item(rw, cl) is not None:
                        valu = self.table.item(rw, cl).text().strip()
                        if len(valu) < 1:
                            continue
                        if self.span is not None and valu == self.span:
                            in_span = True
                        style = dec_fmts[cl]
                        if valu[-1] == '%':
                            is_pct = True
                            i = valu.rfind('.')
                            if i >= 0:
                                dec_pts = (len(valu) - i - 2)
                                style = '#,##0.' + '0' * dec_pts + '%'
                            else:
                                dec_pts = 0
                                style = '#,##0%'
                        else:
                            is_pct = False
                        if hdr_types[cl] == 'int':
                            try:
                                val1 = valu
                                if is_pct:
                                    val1 = val1.strip('%')
                                val1 = val1.replace(',', '')
                                if is_pct:
                                    valu = round(int(val1) / 100., dec_pts + 2)
                                else:
                                    valu = int(val1)
                            except:
                                pass
                        elif hdr_types[cl] == 'float':
                            try:
                                val1 = valu
                                if is_pct:
                                    val1 = val1.strip('%')
                                val1 = val1.replace(',', '')
                                if is_pct:
                                    valu = round(float(val1) / 100., dec_pts + 2)
                                else:
                                    valu = float(val1)
                            except:
                                pass
                        else:
                            if is_pct:
                                try:
                                    val1 = valu.strip('%')
                                    val1 = val1.replace(',', '')
                                    valu = round(float(val1) / 100., dec_pts + 2)
                                except:
                                    pass
                        if not in_span:
                            if is_pct:
                                plus = 3
                            else:
                                plus = 0
                            xl_lens[cl] = max(xl_lens[cl], len(str(valu)) + plus)
                        ws.cell(row=rw + 2, column=cl + 1).value = valu
                        ws.cell(row=rw + 2, column=cl + 1).font = normal
                        ws.cell(row=rw + 2, column=cl + 1).number_format = style
            for cl in range(self.table.columnCount()):
                ws.column_dimensions[ssCol(cl + 1)].width = xl_lens[cl]
            ws.freeze_panes = 'A2'
            wb.save(data_file)
            wb.close()
        self.savedfile = data_file
        if not self.edit_table:
            self.close()

    def replacetbl(self):
    # https://stackoverflow.com/questions/2827623/how-can-i-create-an-object-and-add-attributes-to-it
        self.replaced = {}
        if self.oclass is not None:
            for rw in range(self.table.rowCount()):
                newdict = {}
                key = self.table.item(rw, 0).text()
                for cl in range(1, self.table.columnCount()):
                    if self.labels[self.fields[cl]] == 'int' or self.labels[self.fields[cl]] == 'float':
                        try:
                            tst = self.table.item(rw, cl).text().strip().replace(',', '')
                            if tst == '':
                                tst = '0'
                            if self.labels[self.fields[cl]] == 'int':
                                valu = int(tst)
                            else:
                                valu = float(tst)
                        except:
                            valu = ''
                    else:
                        valu = self.table.item(rw, cl).text()
                    setattr(self.objects[rw], self.fields[cl], valu)
                    newdict[self.fields[cl]] = valu
                self.replaced[key] = newdict
            self.close()
            return
        for rw in range(self.table.rowCount()):
            key = self.table.item(rw, 0).text()
            values = []
            for cl in range(1, self.table.columnCount()):
                if self.table.item(rw, cl) is not None:
                    if self.labels[self.fields[cl]] == 'int' or self.labels[self.fields[cl]] == 'float':
                        tst = self.table.item(rw, cl).text().strip().replace(',', '')
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
                                    elif self.abbr and valu > 99 and valu < 100000:
                                        if len(str(valu / pow(10, 3)).split('.')[1]) < 3:
                                            valu = str(valu / pow(10, 3)) + 'K'
                                        else:
                                            value = str(valu)
                                    elif self.abbr and valu >= 100000:
                                        if len(str(valu / pow(10, 6)).split('.')[1]) < 3:
                                            valu = str(valu / pow(10, 6)) + 'M'
                                        else:
                                            valu = str(valu)
                                    else:
                                        valu = str(valu)
                                except:
                                    if self.txt_ok is not None and self.table.item(rw, 0).text() in self.txt_ok:
                                        valu = self.table.item(rw, cl).text()
                                    else:
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
                                    elif self.abbr and valu > 99 and valu < 1000000:
                                        if len(str(valu / pow(10, 3)).split('.')[1]) < 3:
                                            valu = str(valu / pow(10, 3)) + 'K'
                                            valu = valu.replace('.0K', 'K')
                                        else:
                                            value = str(valu)
                                    elif self.abbr and valu >= 1000000:
                                        if len(str(valu / pow(10, 6)).split('.')[1]) < 3:
                                            valu = str(valu / pow(10, 6)) + 'M'
                                            valu = valu.replace('.0M', 'M')
                                        else:
                                            valu = str(valu)
                                    else:
                                        valu = str(valu)
                                except:
                                    if self.txt_ok is not None and self.table.item(rw, 0).text() in self.txt_ok:
                                        valu = self.table.item(rw, cl).text()
                                    else:
                                        self.message.setText('Error with ' + self.fields[cl].title() + ' field - ' + tst)
                                        self.replace = None
                                        return
                    else:
                        valu = self.table.item(rw, cl).text()
                else:
                    valu = ''
                    try:
                        if isinstance(self.table.cellWidget(rw, cl), QtWidgets.QComboBox):
                            valu = self.table.cellWidget(rw, cl).currentText()
                    except:
                        pass
                if valu != '':
                    values.append(self.fields[cl] + '=' + valu)
            self.replaced[key] = values
        self.close()

    def getValues(self):
        try:
            if self.edit_table:
                return self.replaced
        except:
            pass
        return None

    def getItem(self, col):
        try:
            return self.table.item(self.table.currentRow(), col).text()
        except:
            pass
        return None
