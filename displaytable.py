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

from datetime import datetime
from displaytablebase import TableBase
import displayobject
import openpyxl as oxl
from PyQt5 import QtWidgets, QtGui, QtCore
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

from PyQt5 import QtCore, QtWidgets, QtGui
from displaytablebase import TableBase
from sirenicons import Icons

class Table(QtWidgets.QDialog):
    """
    PyQt5 Table UI class inheriting from TableBase for data handling.
    """

    def __init__(self, objects, fields=None, sumby=None, sumfields=None, save_folder='', units='', decpts=None, abbr=True,
                 title=None, parent=None, **kwargs):
        parent = kwargs.pop('parent', None)
        super().__init__(parent)
        # Pass essential arguments to TableBase
        self.base = TableBase(
            objects=objects,
            fields=fields,
            sumby=sumby,
            sumfields=sumfields,
            units=units,
            decpts=decpts,
            abbr=abbr
        )
        self.icons = Icons()
        self.save_folder = save_folder
        self.setup_ui()

    def setup_ui(self):
        """Setup PyQt UI components using TableBase data."""
        self.setWindowTitle(f'SIREN - Display {self.base.fields[0] if self.base.fields else "Items"}')
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))

        # Message label
        self.message = QtWidgets.QLabel('(Right click column header to sort)')
        
        # Table setup
        self.table = QtWidgets.QTableWidget()
        self.populate_table()

        # Buttons
        self.quitButton = QtWidgets.QPushButton('&Quit')
        self.quitButton.clicked.connect(self.quit)
        self.saveButton = QtWidgets.QPushButton('&Save')
        self.saveButton.clicked.connect(self.saveit)
        
        buttonLayout = QtWidgets.QHBoxLayout()
        buttonLayout.addWidget(self.quitButton)
        buttonLayout.addWidget(self.saveButton)
        buttons = QtWidgets.QFrame()
        buttons.setLayout(buttonLayout)

        # Layout
        layout = QtWidgets.QVBoxLayout()
        layout.addWidget(self.table)
        layout.addWidget(self.message)
        layout.addWidget(buttons)
        self.setLayout(layout)

        self.resize_table()

    def populate_table(self):
        self.labels = {}
        self.lens = {}
        """Populate the QTableWidget with data from TableBase."""
        data = self.base.get_formatted_data()
        fields = self.base.fields
        
        self.table.setRowCount(len(data))
        self.table.setColumnCount(len(fields))
        self.table.setHorizontalHeaderLabels([self.base.nice(f) for f in fields])
        
        for row_idx, row_data in enumerate(data):
            for col_idx, field in enumerate(fields):
                value = str(row_data.get(field, ''))
                item = QtWidgets.QTableWidgetItem(value)
                self.table.setItem(row_idx, col_idx, item)
        
        self.table.horizontalHeader().sectionClicked.connect(self.header_click)

    def header_click(self, index):
        """Handle sorting when header is clicked."""
        field = self.base.fields[index]
        self.base.order_by_field(field)
        self.populate_table()

    def saveit(self):
        if self.title is None:
            iam = getattr(self.objects[0], '__module__')
        else:
            iam = self.title
        data_file = '%s_Table_%s%s.xlsx' % (iam, self.year,
                    datetime.now().strftime('_%Y-%M-%d_%H%M'))
        data_file = QtWidgets.QFileDialog.getSaveFileName(None, 'Save ' + iam + ' Table',
                    self.save_folder + data_file, 'Excel Files (*.xls*);;CSV Files (*.csv)')[0]
        if data_file == '':
            return
        self.base.save_as(data_file, self.table)

    def resize_table(self):
        self.table.resizeColumnsToContents()
        width = sum([self.table.columnWidth(i) for i in range(self.table.columnCount())]) + 50
        height = self.table.rowHeight(0) * (self.table.rowCount() + 4)
        self.resize(width, height)

    def quit(self):
        self.close()
