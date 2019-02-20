#!/usr/bin/python
#
#  Copyright (C) 2016-2019 Sustainable Energy Now Inc., Angus King
#
#  floatstatus.py - This file is part of SIREN.
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
import os
import ssc
import sys
import ConfigParser   # decode .ini file
from PyQt4 import QtGui, QtCore

from editini import SaveIni


class FloatStatus(QtGui.QDialog):
    procStart = QtCore.pyqtSignal(str)

    def __init__(self, mainwindow, scenarios_folder, scenarios):
        super(FloatStatus, self).__init__()
        self.mainwindow = mainwindow
        self.scenarios_folder = scenarios_folder
        self.scenarios = scenarios
        config = ConfigParser.RawConfigParser()
        if len(sys.argv) > 1:
            self.config_file = sys.argv[1]
        else:
            self.config_file = 'SIREN.ini'
        config.read(self.config_file)
        self.restorewindows = False
        self.logged = False
        try:
            rw = config.get('Windows', 'restorewindows')
            if rw.lower() in ['true', 'yes', 'on']:
                self.restorewindows = True
        except:
            pass
        self.log_status = True
        try:
            rw = config.get('Windows', 'log_status')
            if rw.lower() in ['false', 'no', 'off']:
                self.log_status = False
        except:
            pass
        self.be_open = True
        max_line = 0
        lines1 = ''
        max_line = 0
        line_cnt1 = 1
        for line in self.scenarios:
            lines1 += line[0] + '\n'
            line_cnt1 += 1
            if len(line[0]) > max_line:
                max_line = len(line[0])
        if self.log_status or not self.log_status:
            ssc_api = ssc.API()
            comment = 'SAM SDK Core: Version = %s. Build info = %s.' % (ssc_api.version(), ssc_api.build_info())
            lines2 = '%s. SIREN log started\n          Preference File: %s\n          Working directory: %s\n          %s' \
                      % (str(QtCore.QDateTime.toString(QtCore.QDateTime.currentDateTime(), 'hh:mm:ss')), \
                      self.config_file, os.getcwd(), comment)
            line_cnt2 = 1
            if len(lines2) > max_line:
                max_line = len(lines2)
        else:
            line_cnt2 = 0
        self.saveButton = QtGui.QPushButton(self.tr('Save Log'))
        self.saveButton.clicked.connect(self.save)
        self.quitButton = QtGui.QPushButton('Quit')
        self.quitButton.clicked.connect(self.close)
        QtGui.QShortcut(QtGui.QKeySequence('q'), self, self.close)
        buttonLayout = QtGui.QHBoxLayout()
        buttonLayout.addStretch(1)
        buttonLayout.addWidget(self.quitButton)
        buttonLayout.addWidget(self.saveButton)
        self.scenarios = QtGui.QPlainTextEdit()
        self.scenarios.setFont(QtGui.QFont('Courier New', 10))
        fnt = self.scenarios.fontMetrics()
        self.scenarios.setSizePolicy(QtGui.QSizePolicy(QtGui.QSizePolicy.Expanding,
            QtGui.QSizePolicy.Expanding))
        self.scenarios.setPlainText(lines1)
        self.scenarios.setReadOnly(True)
        screen = QtGui.QDesktopWidget().availableGeometry()
        ln = (max_line + 5) * fnt.maxWidth()
        if ln > screen.width() * .80:
            ln = int(screen.width() * .80)
        h1 = (line_cnt1 + 1) * fnt.height()
        if self.log_status:
            self.loglines = QtGui.QPlainTextEdit()
            self.loglines.setFont(QtGui.QFont('Courier New', 10))
            h2 = (line_cnt2 + 2) * fnt.height()
            if h1 + h2 > screen.height() * .80:
       #      h1 = max(int(screen.height() * float(h1 / h2)), int(fnt.height()))
                h2 = int(screen.height() * .80) - h1
            self.loglines.setSizePolicy(QtGui.QSizePolicy(QtGui.QSizePolicy.Expanding,
                QtGui.QSizePolicy.Expanding))
            self.loglines.setPlainText(lines2)
            self.loglines.setReadOnly(True)
            self.loglines.resize(ln, h2)
        self.scenarios.resize(ln, h1)
        self.scenarios.setFixedHeight(h1)
        layout = QtGui.QVBoxLayout()
        layout.addWidget(QtGui.QLabel('Open scenarios'))
        layout.addWidget(self.scenarios)
        if self.log_status:
            layout.addWidget(QtGui.QLabel('Session log'))
            layout.addWidget(self.loglines)
        layout.addLayout(buttonLayout)
        self.setLayout(layout)
        self.setWindowTitle('SIREN - Status - ' + self.config_file)
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        if self.restorewindows:
            try:
                rw = config.get('Windows', 'log_size').split(',')
                self.resize(int(rw[0]), int(rw[1]))
                mp = config.get('Windows', 'log_pos').split(',')
                self.move(int(mp[0]), int(mp[1]))
            except:
                pass
        self.show()

    def save(self):
        save_filename = self.scenarios_folder + 'SIREN_Log_' + self.config_file[:-4]
        save_filename += '_' + str(QtCore.QDateTime.toString(QtCore.QDateTime.currentDateTime(), 'yyyy-MM-dd_hhmm'))
        save_filename += '.txt'
        fileName = QtGui.QFileDialog.getSaveFileName(self, self.tr("QFileDialog.getSaveFileName()"),
                   save_filename, self.tr("All Files (*);;Text Files (*.txt)"))
         # save scenarios list and log
        if not fileName.isEmpty():
            s = open(fileName, 'w')
            s.write('Scenarios:\n')
            t = str(self.scenarios.toPlainText())
            bits = t.split('\n')
            for lin in bits:
                if len(lin) > 0:
                    s.write(' ' * 10 + lin + '\n')
            if self.log_status:
                s.write('\nSession log:\n\n')
                t = str(self.loglines.toPlainText())
                bits = t.split('\n')
                for lin in bits:
                    if len(lin) > 0:
                        s.write(lin + '\n')
            s.close()
            self.logged = False

    @QtCore.pyqtSlot()
    def exit(self):
        if self.log_status:
            self.log('SIREN log stopped')
        self.be_open = False
        self.close()

    @QtCore.pyqtSlot()
    def log(self, text):
        self.loglines.appendPlainText(str(QtCore.QDateTime.toString(QtCore.QDateTime.currentDateTime(),
                                'hh:mm:ss. ')) + text)
        self.logged = True

    @QtCore.pyqtSlot()
    def log2(self, text):
        self.loglines.appendPlainText(text)
        self.logged = True

    @QtCore.pyqtSlot()
    def updateScenarios(self, scenarios):
        lines1 = ''
        for line in scenarios:
            lines1 += line[0] + '\n'
        self.scenarios.setPlainText(lines1)

    def closeEvent(self, event):
        if self.logged:
            reply = QtGui.QMessageBox.question(self, 'SIREN Status',
                    'Do you want to save Session log?', QtGui.QMessageBox.Yes |
                    QtGui.QMessageBox.No, QtGui.QMessageBox.No)
            if reply == QtGui.QMessageBox.Yes:
                self.save()
        if self.restorewindows:
            updates = {}
            lines = []
            add = int((self.frameSize().width() - self.size().width()) / 2)   # need to account for border
            lines.append('log_pos=%s,%s' % (str(self.pos().x() + add), str(self.pos().y() + add)))
            lines.append('log_size=%s,%s' % (str(self.width()), str(self.height())))
            updates['Windows'] = lines
            SaveIni(updates)
        if self.be_open:
            self.procStart.emit('goodbye')
        event.accept()
