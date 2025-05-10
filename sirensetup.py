#!/usr/bin/python3
#
#  Copyright (C) 2024 Sustainable Energy Now Inc., Angus King
#
#  sirensetup.py - This file is part of SIREN.
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
from PyQt5 import QtCore, QtGui, QtWidgets
from shutil import copy2
import subprocess
import sys
from credits import fileVersion
import displayobject
from getmodels import commonprefix
from editini import EdtDialog, SaveIni
from senutils import ClickableQLabel


class SetUp(QtWidgets.QWidget):
    statusmsg = QtCore.pyqtSignal()

    def __init__(self, help='help.html'):
        super(SetUp, self).__init__()
        self.help = help
        curdir = os.getcwd()
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            fldr_div = '\\'
        else:
            fldr_div = '/'
        mydir = os.path.abspath(__file__)
        updir = mydir[:mydir.rfind(fldr_div) + 1]
        self.grid = QtWidgets.QGridLayout()
        self.grid.addWidget(QtWidgets.QLabel('Setup:'), 0, 0)
        msg = 'SIREN requires an area of storage that it can write to.\n' + \
              'Use this program to copy (both) the SIREN model and\n' + \
              'Sample Scenario files to a writeable location (e.g. OneDrive).\n\n' + \
              "Then run 'run_siren.bat' from the Model folder to run SIREN."
        self.grid.addWidget(QtWidgets.QLabel(msg), 0, 1, 3, 4)
        rw = 4
        samples = ''
        self.msi = ''
        if os.path.exists(updir + 'siren_sample'):
            samples = updir + 'siren_sample'
        elif sys.platform == 'win32' or sys.platform == 'cygwin': # maybe in ProgramData
            if os.path.exists('C:\\ProgramData\\SIREN\\siren_sample'):
                samples = 'C:\\ProgramData\\SIREN\\siren_sample'
                self.msi = 'msi'
            else:
                up1more = updir[:updir[:-1].rfind(fldr_div) + 1]
                up2more = up1more[:up1more[:-1].rfind(fldr_div) + 1]
                up3more = up2more[:up2more[:-1].rfind(fldr_div) + 1]
                if os.path.exists(f'{up3more}\\ProgramData\\SIREN\\siren_sample'):
                    samples = f'{up3more}\\ProgramData\\SIREN\\siren_sample'
                    self.msi = 'msi'
        elif os.path.exists(curdir + '/windows/dist/siren_sample'):
            samples = curdir + '/windows/dist/siren_sample'
        self.dir_labels = ['Sample', 'Model', 'Scenario']
        self.dirs = []
        for dirn in self.dir_labels:
           self.grid.addWidget(QtWidgets.QLabel(f'{dirn} files location:'), rw, 0)
           self.dirs.append(ClickableQLabel())
           if dirn == 'Sample' and samples != '':
               self.dirs[-1].setStyleSheet("border: 1px inset grey; min-height: 22px; border-radius: 4px;")
               self.dirs[-1].setText(samples)
           else:
               self.dirs[-1].setStyleSheet("background-color: white; border: 1px inset grey; min-height: 22px; border-radius: 4px;")
               if dirn != 'Scenario':
                   self.dirs[-1].setText(curdir)
               self.dirs[-1].clicked.connect(self.dirChanged)
           self.grid.addWidget(self.dirs[-1], rw, 1, 1, 5)
           rw += 1
        self.grid.addWidget(QtWidgets.QLabel('Replace files:'), rw, 0)
        msg = '(check to replace existing files)'
        self.replace = QtWidgets.QCheckBox(msg, self)
        self.replace.setCheckState(QtCore.Qt.Unchecked)
        self.grid.addWidget(self.replace, rw, 1)
        rw += 1
        self.log = QtWidgets.QLabel('')
        msg_palette = QtGui.QPalette()
        msg_palette.setColor(QtGui.QPalette.Foreground, QtCore.Qt.red)
        self.log.setPalette(msg_palette)
        self.grid.addWidget(self.log, rw, 0, 1, 5)
        rw += 1
        quit = QtWidgets.QPushButton('Quit', self)
        self.grid.addWidget(quit, rw, 0)
        quit.clicked.connect(self.quitClicked)
        QtWidgets.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
        docopy = QtWidgets.QPushButton('Copy files', self)
        self.grid.addWidget(docopy, rw, 1)
        docopy.clicked.connect(self.docopyClicked)
        doedit = QtWidgets.QPushButton('Edit Model file', self)
        self.grid.addWidget(doedit, rw, 2)
        doedit.clicked.connect(self.editIniFile)
        launch = QtWidgets.QPushButton('Launch SIREN', self)
        self.grid.addWidget(launch, rw, 3)
        launch.clicked.connect(self.spawn)
        help = QtWidgets.QPushButton('Help', self)
        self.grid.addWidget(help, rw, 4)
        help.clicked.connect(self.helpClicked)
        QtWidgets.QShortcut(QtGui.QKeySequence('F1'), self, self.helpClicked)
        frame = QtWidgets.QFrame()
        frame.setLayout(self.grid)
        self.scroll = QtWidgets.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtWidgets.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
  #      self.resize(self.width() + int(self.world_width * .7), self.height() + int(self.world_height * .7))
        self.setWindowTitle('SIREN - sirensetup (' + fileVersion() + ') - Choose Models and Scenarios Locations')
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        self.center()
        self.resize(int(self.sizeHint().width() * 1.5), int(self.sizeHint().height() * 1.3))
        self.show()

    def center(self):
        frameGm = self.frameGeometry()
        screen = QtWidgets.QApplication.desktop().screenNumber(QtWidgets.QApplication.desktop().cursor().pos())
        centerPoint = QtWidgets.QApplication.desktop().availableGeometry(screen).center()
        frameGm.moveCenter(centerPoint)
        self.move(frameGm.topLeft())

    def zoomChanged(self, val):
        self.zoomScale.setText('(' + scale[int(val)] + ')')
        self.zoomScale.adjustSize()

    def dirChanged(self):
        for i in range(len(self.dir_labels)):
            if self.dirs[i].hasFocus():
                break
        curdir = self.dirs[i].text()
        newdir = QtWidgets.QFileDialog.getExistingDirectory(self, f'Choose {self.dir_labels[i]} Folder',
                 curdir, QtWidgets.QFileDialog.ShowDirsOnly)
        if newdir != '':
            self.dirs[i].setText(newdir)

    def docopyClicked(self):
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            fldr_div = '\\'
        else:
            fldr_div = '/'
        if self.dirs[2].text() == '':
            self.dirs[2].setText(self.dirs[1].text())
        for i in range(len(self.dir_labels)):
            if not os.path.exists(self.dirs[i].text()):
                self.log.setText(f'{self.dir_labels[i]} folder not found - {self.dirs[i].text()}')
                return
        mydir = os.path.abspath(__file__)
        updir = mydir[:mydir.rfind(fldr_div) + 1]
        # .ini files first
        ini_count = 0
        if self.replace.isChecked() or not os.path.exists(f'{self.dirs[1].text()}{fldr_div}getfiles.ini'):
            try:
                if os.path.exists(f'{self.dirs[0].text()}{fldr_div}getfiles.ini'):
                    copy2(f'{self.dirs[0].text()}{fldr_div}getfiles.ini', f'{self.dirs[1].text()}{fldr_div}getfiles.ini')
                else:
                    copy2(f'{updir}{fldr_div}getfiles.ini', f'{self.dirs[1].text()}{fldr_div}getfiles.ini')
                ini_count += 1
            except:
                pass
        if self.replace.isChecked() or not os.path.exists(f'{self.dirs[1].text()}{fldr_div}SIREN.ini'):
            try:
                copy2(f'{self.dirs[0].text()}{fldr_div}SIREN{self.msi}.ini', f'{self.dirs[1].text()}{fldr_div}SIREN.ini')
                ini_count += 1
                # update scenarios property
                newfldr = f'{self.dirs[2].text()}'
                that_len = len(commonprefix([updir, newfldr]))
                if that_len > 0:
                    bits = updir[that_len:].split(fldr_div)
                    pfx = ('..' + fldr_div) * (len(bits) - 1)
                    newfldr = pfx + newfldr[that_len + 1:]
                updates = {'Files': [f'scenarios={newfldr}']}
                SaveIni(updates, ini_file=f'{self.dirs[1].text()}{fldr_div}SIREN.ini')
            except:
                pass
        bat_file = ''
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            bat_file += f'{mydir[:2]}\n'
            bat_file += f'cd "{updir}"\n'
            bat_file += f'siren.exe "{self.dirs[1].text()}"\n'
        else:
            bat_file += f'cd "{updir}"\n'
            bat_file += f'python siren.py "{self.dirs[1].text()}"\n'
        bf = open(f'{self.dirs[1].text()}{fldr_div}run_siren.bat', 'w')
        bf.write(bat_file)
        bf.close()
        # copy files from my_scenarios
        scenario_count = 0
        the_files = os.listdir(f'{self.dirs[0].text()}{fldr_div}my_scenarios')
        for a_file in the_files:
            if self.replace.isChecked() or not os.path.exists(f'{self.dirs[2].text()}{fldr_div}{a_file}'):
                copy2(f'{self.dirs[0].text()}{fldr_div}my_scenarios{fldr_div}{a_file}',
                      f'{self.dirs[2].text()}{fldr_div}{a_file}')
                scenario_count += 1
        self.log.setText(f'{ini_count} Preferences (Models) files and {scenario_count} Scenario files copied')
        # if os.path.exists(dir1 + 'SIREN.ini'):
        #     if os.path.exists(dir1 + 'SIREN.ini~'):
        #         os.remove(dir1 + 'SIREN.ini~')
        #     os.rename(dir1 + 'SIREN.ini', dir1 + 'SIREN.ini~')
        return

    def editIniFile(self):
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            fldr_div = '\\'
        else:
            fldr_div = '/'
        config_file = f'{self.dirs[1].text()}{fldr_div}SIREN.ini'
        dialr = EdtDialog(config_file)
        dialr.exec_()

    def spawn(self):
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            fldr_div = '\\'
            if not os.path.exists(f'{self.dirs[1].text()}{fldr_div}run_siren.bat'):
                return
            bat_file = f'{self.dirs[1].text()}{fldr_div}run_siren.bat'.replace('/', fldr_div)
            pid = subprocess.Popen([bat_file]).pid
        else:
            if not os.path.exists(f'{self.dirs[1].text()}/run_siren.bat'):
                return
            pid = subprocess.Popen(['sh', f'{self.dirs[1].text()}/./run_siren.bat']).pid
        self.log.setText('SIREN launched')
        return

    def helpClicked(self):
        dialog = displayobject.AnObject(QtWidgets.QDialog(), self.help,
                 title='Help for installation (' + fileVersion() + ')', section='install')
        dialog.exec_()

    def quitClicked(self):
        self.close()

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    ex = SetUp()
    app.exec_()
    app.deleteLater()
    sys.exit()
