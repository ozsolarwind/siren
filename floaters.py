#!/usr/bin/python3
#
#  Copyright (C) 2016-2020 Sustainable Energy Now Inc., Angus King
#
#  floaters.py - This file is part of SIREN.
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
import configparser   # decode .ini file
from PyQt4 import QtGui, QtCore
from editini import SaveIni
from getmodels import getModelFile


class FloatLegend(QtGui.QDialog):
    procStart = QtCore.pyqtSignal(str)

    def __init__(self, techdata, stations, flags):
        super(FloatLegend, self).__init__()
        self.stations = stations
        self.techdata = techdata
        self.flags = flags
        self.initUI()

    def initUI(self):
        config = configparser.RawConfigParser()
        if len(sys.argv) > 1:
            config_file = sys.argv[1]
        else:
            config_file = getModelFile('SIREN.ini')
        config.read(config_file)
        self.restorewindows = False
        try:
            rw = config.get('Windows', 'restorewindows')
            if rw.lower() in ['true', 'yes', 'on']:
                self.restorewindows = True
        except:
            pass
        self.be_open = True
        self.grid = QtGui.QGridLayout()
        tech_sizes = {}
        txt_lens = [10, 0]
        for st in self.stations:
            if st.technology[:6] == 'Fossil' and not self.flags[3]:
                continue
            if st.technology not in tech_sizes:
                tech_sizes[st.technology] = 0.
            if st.technology == 'Wind':
                tech_sizes[st.technology] += self.techdata[st.technology][0] * float(st.no_turbines) * pow((st.rotor * .001), 2)
            else:
                tech_sizes[st.technology] += self.techdata[st.technology][0] * float(st.capacity)
            txt_lens[0] = max(txt_lens[0], len(st.technology))
        total_area = 0.
        row = 0
        for key, value in iter(sorted(tech_sizes.items())):
            labl = QtGui.QLabel('__')
            colr = QtGui.QColor(self.techdata[key][1])
            labl.setStyleSheet('QLabel {background-color: %s; color: %s;}' % (colr.name(), colr.name()))
            self.grid.addWidget(labl, row, 0)
            self.grid.addWidget(QtGui.QLabel(key), row, 1)
            total_area += value
            area = '%s sq. Km' % '{:0.1f}'.format(value)
            txt_lens[1] = max(txt_lens[1], len(area))
            self.grid.addWidget(QtGui.QLabel(area), row, 2)
            self.grid.setRowStretch(row, 0)
            row += 1
        self.grid.setColumnStretch(0, 0)
        self.grid.setColumnStretch(1, 1)
        self.grid.addWidget(QtGui.QLabel('Total Area'), row, 1)
        area = '%s sq. Km' % '{:0.1f}'.format(total_area)
        txt_lens[1] = max(txt_lens[1], len(area))
        self.grid.addWidget(QtGui.QLabel(area), row, 2)
        self.grid.setRowStretch(row, 1)
        row += 1
        if self.flags[0]:
            txt = 'Circles show relative capacity in MW'
        elif self.flags[1]:
            txt = 'Circles show relative generation in MWh'
        else:
            txt = 'Station '
            if self.flags[2]:
                txt += 'Squares'
            else:
                txt += 'Circles'
            txt += ' show estimated area in sq. Km'
        if len(txt) > (txt_lens[0] + txt_lens[1]):
            txts = txt.split(' ')
            txt = txts[0]
            ln = len(txt)
            for i in range(1, len(txts)):
                if ln + len(txts[i]) > (txt_lens[0] + txt_lens[1]):
                    txt += '\n'
                    txt += txts[i]
                    ln = len(txts[i])
                else:
                    txt += ' ' + txts[i]
                    ln += len(txts[i]) + 1
        self.grid.addWidget(QtGui.QLabel(txt), row, 1, 1, 2)
        self.grid.setRowStretch(row, 1)
        row += 1
        quit = QtGui.QPushButton('Quit', self)
        self.grid.addWidget(quit, row, 1)
        self.grid.setRowStretch(row, 1)
        self.grid.setVerticalSpacing(10)
        quit.clicked.connect(self.quitClicked)
        QtGui.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
        frame = QtGui.QFrame()
        frame.setLayout(self.grid)
        self.scroll = QtGui.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtGui.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
        self.setWindowTitle('SIREN - Legend')
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        screen = QtGui.QApplication.desktop().primaryScreen()
        scr_right = QtGui.QApplication.desktop().availableGeometry(screen).right()
        scr_bottom = QtGui.QApplication.desktop().availableGeometry(screen).bottom()
        win_width = self.sizeHint().width()
        win_height = self.sizeHint().height()
        if self.restorewindows:
            try:
                rw = config.get('Windows', 'legend_size').split(',')
                lst_width = int(rw[0])
                lst_height = int(rw[1])
                mp = config.get('Windows', 'legend_pos').split(',')
                lst_left = int(mp[0])
                lst_top = int(mp[1])
                lst_right = lst_left + lst_width
                lst_bottom = lst_top + lst_height
                screen = QtGui.QApplication.desktop().screenNumber(QtCore.QPoint(lst_left, lst_top))
                scr_right = QtGui.QApplication.desktop().availableGeometry(screen).right()
                scr_left = QtGui.QApplication.desktop().availableGeometry(screen).left()
                if lst_right < scr_right:
                    if (lst_right - win_width) >= scr_left:
                        scr_right = lst_right
                    else:
                        scr_right = scr_left + win_width
                scr_bottom = QtGui.QApplication.desktop().availableGeometry(screen).bottom()
                scr_top = QtGui.QApplication.desktop().availableGeometry(screen).top()
                if lst_bottom < scr_bottom:
                    if (lst_bottom - win_height) >= scr_top:
                        scr_bottom = lst_bottom
                    else:
                        scr_bottom = scr_top + win_height
            except:
                pass
        win_left = scr_right - win_width
        win_top = scr_bottom - win_height
        self.resize(win_width, win_height)
        self.move(win_left, win_top)
        self.show()

    @QtCore.pyqtSlot()
    def exit(self):
        self.be_open = False
        self.close()

    def closeEvent(self, event):
        if self.restorewindows:
            updates = {}
            lines = []
            add = int((self.frameSize().width() - self.size().width()) / 2)   # need to account for border
            lines.append('legend_pos=%s,%s' % (str(self.pos().x() + add), str(self.pos().y() + add)))
            lines.append('legend_size=%s,%s' % (str(self.width()), str(self.height())))
            updates['Windows'] = lines
            SaveIni(updates)
        if self.be_open:
            self.procStart.emit('goodbye')
        event.accept()

    def quitClicked(self):
        self.close()


class FloatMenu(QtGui.QDialog):
    procStart = QtCore.pyqtSignal(str)
    procAction = QtCore.pyqtSignal(QtGui.QAction)

    def __init__(self, menubar):
        super(FloatMenu, self).__init__()
        self.menubar = menubar
        self.initUI()

    def initUI(self):
        config = configparser.RawConfigParser()
        if len(sys.argv) > 1:
            config_file = sys.argv[1]
        else:
            config_file = getModelFile('SIREN.ini')
        config.read(config_file)
        self.restorewindows = False
        try:
            rw = config.get('Windows', 'restorewindows')
            if rw.lower() in ['true', 'yes', 'on']:
                self.restorewindows = True
        except:
            pass
        try:
            open_menus = config.get('Windows', 'menu_open').split(',')
        except:
            open_menus = ''
        self.be_open = True
        self.menus = []
        for lvl1 in self.menubar.actions():
            if str(lvl1.text()).find('&') >= 0:
                try:
                    self.menus.append([str(lvl1.text()).replace('&', ''), lvl1.icon(), None, []])
                except:
                    self.menus.append([str(lvl1.text()).replace('&', ''), None, None, []])
                try:
                    for lvl2 in lvl1.menu().actions():
                        if str(lvl2.text()).find('&') >= 0:
                            try:
                                self.menus[-1][-1].append([str(lvl2.text()).replace('&', ''), lvl2.icon(), None, []])
                            except:
                                self.menus[-1][-1].append([str(lvl2.text()).replace('&', ''), None, None, []])
                            try:
                                for lvl3 in lvl2.menu().actions():
                                    self.menus[-1][-1][-1][-1].append([str(lvl3.text()),
                                        lvl3.icon(), lvl3, '3'])
                            except:
                                pass
                        else:
                            self.thisaction = lvl2
                            self.menus[-1][-1].append([str(lvl2.text()), lvl2.icon(), lvl2, '2'])
                except:
                    pass
            else:
                self.menus.append([str(lvl1.text()), lvl1.icon(), lvl1.actions()[0], '1'])
        self.grid = QtGui.QGridLayout()
        ctr = 0
        self.butn = []
        self.topmenus = {}
        self.buttons = {}
        for i in range(len(self.menus)):
            self.butn.append(QtGui.QPushButton(self.menus[i][0], self))
            self.butn[-1].setIcon(self.menus[i][1])
            self.butn[-1].setStyleSheet('QPushButton {color: #005fb6; border: 2px solid #e65900;' +
                               ' border-radius: 6px;}')
            self.butn[-1].clicked.connect(self.menuClicked)
            self.grid.addWidget(self.butn[-1], ctr, 0, 1, 3)
            ctr += 1
            self.topmenus[self.menus[i][0]] = [False, ctr, 0]
            if type(self.menus[i][3]) is list:
                for j in range(len(self.menus[i][3])):
                    self.butn.append(QtGui.QPushButton(self.menus[i][3][j][0], self))
                    self.butn[-1].setIcon(self.menus[i][3][j][1])
                #     self.butn[-1].setStyleSheet('QPushButton {Text-align:left;}')
                    self.butn[-1].clicked.connect(self.buttonClicked)
                    self.butn[-1].hide()
                    self.buttons[self.menus[i][3][j][0]] = self.menus[i][3][j][2]
                    self.grid.addWidget(self.butn[-1], ctr, 1, 1, 2)
                    ctr += 1
                    if type(self.menus[i][3][j][3]) is list:
                        for k in range(len(self.menus[i][3][j][3])):
                            self.butn.append(QtGui.QPushButton(self.menus[i][3][j][3][k][0], self))
                            self.butn[-1].setIcon(self.menus[i][3][j][3][k][1])
                          #   self.butn[-1].setStyleSheet('QPushButton {Text-align:left;}')
                            self.butn[-1].clicked.connect(self.buttonClicked)
                            self.butn[-1].hide()
                            self.buttons[self.menus[i][3][j][3][k][0]] = self.menus[i][3][j][3][k][2]
                            self.grid.addWidget(self.butn[-1], ctr, 2)
                            ctr += 1
            self.topmenus[self.menus[i][0]][2] = ctr
        quit = QtGui.QPushButton('Quit Menu', self)
        self.grid.addWidget(quit, ctr, 0)
        quit.clicked.connect(self.close)
        QtGui.QShortcut(QtGui.QKeySequence('q'), self, self.close)
        for mnu in open_menus:
            strt = self.topmenus[mnu][1]
            stop = self.topmenus[mnu][2]
            for j in range(strt, stop):
                self.butn[j].show()
            self.topmenus[mnu][0] = True
        frame = QtGui.QFrame()
        frame.setLayout(self.grid)
        self.scroll = QtGui.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtGui.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
        self.setWindowTitle('SIREN - Menu')
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        if self.restorewindows:
            try:
                rw = config.get('Windows', 'menu_size').split(',')
                self.resize(int(rw[0]), int(rw[1]))
                mp = config.get('Windows', 'menu_pos').split(',')
                self.move(int(mp[0]), int(mp[1]))
            except:
                pass
        self.show()

    def menuClicked(self, event):
        sender = self.sender()
        strt = self.topmenus[str(sender.text())][1]
        stop = self.topmenus[str(sender.text())][2]
        if self.topmenus[str(sender.text())][0]:
            for j in range(strt, stop):
                self.butn[j].hide()
            self.topmenus[str(sender.text())][0] = False
        else:
            for j in range(strt, stop):
                self.butn[j].show()
            self.topmenus[str(sender.text())][0] = True
    #     self.topmenus[str(sender.text())][0] = self.topmenus[str(sender.text())][0] != True

    @QtCore.pyqtSlot(QtGui.QAction)
    def buttonClicked(self, event):
        sender = self.sender()
        self.thetext = str(sender.text())
        self.procAction.emit(self.buttons[str(sender.text())])

    @QtCore.pyqtSlot()
    def exit(self):
        self.be_open = False
        self.close()

    def text(self):
        return self.thetext

    def closeEvent(self, event):
        if self.restorewindows:
            updates = {}
            lines = []
            add = int((self.frameSize().width() - self.size().width()) / 2)   # need to account for border
            lines.append('menu_pos=%s,%s' % (str(self.pos().x() + add), str(self.pos().y() + add)))
            lines.append('menu_size=%s,%s' % (str(self.width()), str(self.height())))
            open_menus = ''
            for key in list(self.topmenus.keys()):
                if self.topmenus[key][0]:
                    open_menus += ',' + key
            if open_menus != '':
                lines.append('menu_open=%s' % open_menus[1:])
            updates['Windows'] = lines
            SaveIni(updates)
        if self.be_open:
            self.procStart.emit('goodbye')
        event.accept()


class ProgressBar(QtGui.QDialog):
    procStart = QtCore.pyqtSignal(str)

    def __init__(self, minimum=0, maximum=100, msg=None, title=None):
        super(ProgressBar, self).__init__()
    #    self.mainwindow = mainwindow
        self.log_progress = True
        self.be_open = True
        self.progressbar = QtGui.QProgressBar()
        self.progressbar.setMinimum(minimum)
        self.progressbar.setMaximum(maximum)
        self.progressbar.setStyleSheet('QProgressBar {border: 1px solid grey; border-radius: 2px; text-align: center;}' \
                                       + 'QProgressBar::chunk { background-color: #6891c6;}')
        self.button = QtGui.QPushButton('Stop')
        self.button.clicked.connect(self.stopit)
        if msg == None:
             msg = 'Note: Solar Thermal Stations take a while to process'
        self.progress_msg = QtGui.QLabel(msg)
        main_layout = QtGui.QGridLayout()
        main_layout.addWidget(self.button, 0, 0)
        main_layout.addWidget(self.progressbar, 0, 1)
        main_layout.addWidget(self.progress_msg, 1, 1)
        self.setLayout(main_layout)
        if title == None:
            title = 'SIREN - Power Model Progress'
        self.setWindowTitle(title)
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        self.resize(250, 30)
        self.setVisible(False)
        self.show()

    @QtCore.pyqtSlot()
    def stopit(self):
        self.be_open = False

    @QtCore.pyqtSlot()
    def exit(self):
        self.progress_msg.setText('Stop received')
        self.be_open = False
        self.close()

    @QtCore.pyqtSlot(int, int)
    def range(self, minimum, maximum, msg=None):
        self.progressbar.setMinimum(minimum)
        self.progressbar.setMaximum(maximum)
        if msg == None:
            msg = 'Note: Solar Thermal Stations take a while to process'
        self.progress_msg.setText(msg)
        self.be_open = True
        self.setVisible(True)

    @QtCore.pyqtSlot(int, str)
    def progress(self, ctr, message=''):
        if ctr < 0 or ctr == self.progressbar.maximum():
            self.setVisible(False)
        else:
            if ctr % 2:
                self.progressbar.setStyleSheet('QProgressBar {border: 1px solid grey; border-radius: 2px; text-align: center;}' \
                                             + 'QProgressBar::chunk { background-color: #6891c6;}')
            else:
                self.progressbar.setStyleSheet('QProgressBar {border: 1px solid grey; border-radius: 2px; text-align: center;}' \
                                             + 'QProgressBar::chunk { background-color: #CB6720;}')
            self.progressbar.setValue(ctr)
            self.progress_msg.setText(message)

    def closeEvent(self, event):
        if self.be_open:
            self.procStart.emit('goodbye')
        event.accept()


class FloatStatus(QtGui.QDialog):
    procStart = QtCore.pyqtSignal(str)

    def __init__(self, mainwindow, scenarios_folder, scenarios, program='SIREN'):
        super(FloatStatus, self).__init__()
        self.mainwindow = mainwindow
        self.scenarios_folder = scenarios_folder
        self.scenarios = scenarios
        self.program = program
        if scenarios is None:
            self.full_log = False
        else:
            self.full_log = True
        config = configparser.RawConfigParser()
        if len(sys.argv) > 1:
            self.config_file = sys.argv[1]
        else:
            self.config_file = getModelFile('SIREN.ini')
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
        lines2 = ''
        max_line = 0
        line_cnt1 = 1
        line_cnt2 = 0
        if self.full_log:
            try:
                for line in self.scenarios:
                    lines1 += line[0] + '\n'
                    line_cnt1 += 1
                    if len(line[0]) > max_line:
                        max_line = len(line[0])
                if self.log_status or not self.log_status:
                    ssc_api = ssc.API()
                    comment = 'SAM SDK Core: Version = %s. Build info = %s.' \
                              % (ssc_api.version(), ssc_api.build_info().decode())
                    line = '%s. SIREN log started\n          Preference File: %s' + \
                           '\n          Working directory: %s' + \
                           '\n          Scenarios folder: %s\n          %s'
                    lines2 = line % (str(QtCore.QDateTime.toString(QtCore.QDateTime.currentDateTime(), 'hh:mm:ss')), \
                             self.config_file, os.getcwd(), self.scenarios_folder, comment)
                    line_cnt2 = 1
                    if len(lines2) > max_line:
                        max_line = len(lines2)
                else:
                    line_cnt2 = 0
            except:
                pass
        else:
            line = '%s. %s log started\n          Preference File: %s' + \
                   '\n          Working directory: %s' + \
                   '\n          Scenarios folder: %s'
            lines2 = line % (str(QtCore.QDateTime.toString(QtCore.QDateTime.currentDateTime(), 'hh:mm:ss')), \
                     self.program, self.config_file, os.getcwd(), self.scenarios_folder)
            line_cnt2 = 1
            max_line = len(lines2)
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
        if self.full_log:
            layout.addWidget(QtGui.QLabel('Open scenarios'))
            layout.addWidget(self.scenarios)
        if self.log_status:
            layout.addWidget(QtGui.QLabel('Session log'))
            layout.addWidget(self.loglines)
        layout.addLayout(buttonLayout)
        self.setLayout(layout)
        self.setWindowTitle(self.program + ' - Status - ' + self.config_file)
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
        save_filename = self.scenarios_folder + self.program + '_Log_' + self.config_file[:-4]
        save_filename += '_' + str(QtCore.QDateTime.toString(QtCore.QDateTime.currentDateTime(), 'yyyy-MM-dd_hhmm'))
        save_filename += '.txt'
        fileName = QtGui.QFileDialog.getSaveFileName(self, self.tr("QFileDialog.getSaveFileName()"),
                   save_filename, self.tr("All Files (*);;Text Files (*.txt)"))
         # save scenarios list and log
        if fileName != '':
            if not self.full_log and not self.log_status:
                return
            s = open(fileName, 'w')
            if self.full_log:
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
            self.log(self.program + ' log stopped')
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
            reply = QtGui.QMessageBox.question(self, self.program + ' Status',
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
