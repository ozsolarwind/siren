#!/usr/bin/python
#
#  Copyright (C) 2015-2016 Sustainable Energy Now Inc., Angus King
#
#  samrun.py - This file is part of SIREN.
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

import sys
from PyQt4 import QtCore, QtGui
import subprocess as sp

import ssc   # contains all Python classes for accessing ssc

from senuser import getUser


def spaceSplit(string, dropquote=False):
    last = 0
    splits = []
    inQuote = None
    for i, letter in enumerate(string):
        if inQuote:
            if (letter == inQuote):
                inQuote = None
                if dropquote:
                    splits.append(string[last:i])
                    last = i + 1
                    continue
        else:
            if (letter == '"' or letter == "'"):
                inQuote = letter
                if dropquote:
                    last += 1
        if not inQuote and letter == ' ':
            splits.append(string[last:i])
            last = i + 1
    if last < len(string):
        splits.append(string[last:])
    return splits


class RptDialog(QtGui.QDialog):
    def __init__(self, parms, return_code, output, comment=None, parent=None):
        super(RptDialog, self).__init__()
        self.parms = parms
        if isinstance(output, str):
            self.lines = 'Program:\n'
            line_cnt = 1
            max_line = 9
            if comment is not None:
                self.lines += '    ' + self.parms[0] + '\n'
                self.lines += 'Comment:\n'
                line_cnt += 2
                cmt = comment.split('\n')
                for i in range(len(cmt)):
                    self.lines += '    ' + cmt[i] + '\n'
                    line_cnt += 1
                    if (len(cmt[i]) + 5) > max_line:
                        max_line = len(cmt[i]) + 5
            else:
                for i in range(len(self.parms)):
                    if i == 1:
                        self.lines += 'Parameters:\n'
                        line_cnt += 1
                    self.lines += '    ' + self.parms[i] + '\n'
                    if (len(self.parms[i]) + 5) > max_line:
                        max_line = len(self.parms[i]) + 5
                    line_cnt += 1
            self.lines += return_code
            lin = return_code.split('\n')
            line_cnt += len(lin)
            self.lines += output
            lin = output.split('\n')
            line_cnt += len(lin)
            for i in range(len(lin)):
                if len(lin[i]) > max_line:
                    max_line = len(lin[i])
            del lin
        elif isinstance(output, list):
            self.lines = ''
            max_line = 0
            line_cnt = 0
            for line in output:
                self.lines += line + '\n'
                line_cnt += 1
                if len(line) > max_line:
                    max_line = len(line)
        else:
            if type(output) == 'file':
                files = [output]
            else:
                files = output
            self.lines = ''
            max_line = 0
            line_cnt = 0
            for fil in files:
                for line in fil.readlines():
                    self.lines += line
                    line_cnt += 1
                    if len(line) > max_line:
                        max_line = len(line)
            if len(self.lines) < 1:
                self.close
                return
        QtGui.QDialog.__init__(self, parent)
        self.saveButton = QtGui.QPushButton(self.tr('&Save'))
        self.cancelButton = QtGui.QPushButton(self.tr('Cancel'))
        buttonLayout = QtGui.QHBoxLayout()
        buttonLayout.addStretch(1)
        buttonLayout.addWidget(self.saveButton)
        buttonLayout.addWidget(self.cancelButton)
        self.connect(self.saveButton, QtCore.SIGNAL('clicked()'), self,
                     QtCore.SLOT('accept()'))
        self.connect(self.cancelButton, QtCore.SIGNAL('clicked()'),
                     self, QtCore.SLOT('reject()'))
        self.widget = QtGui.QTextEdit()
        self.widget.setFont(QtGui.QFont('Courier New', 10))
        fnt = self.widget.fontMetrics()
        ln = (max_line + 5) * fnt.maxWidth()
        ln2 = (line_cnt + 2) * fnt.height()
        screen = QtGui.QDesktopWidget().availableGeometry()
        if ln > screen.width() * .67:
            ln = int(screen.width() * .67)
        if ln2 > screen.height() * .67:
            ln2 = int(screen.height() * .67)
        self.widget.setSizePolicy(QtGui.QSizePolicy(QtGui.QSizePolicy.Expanding,
            QtGui.QSizePolicy.Expanding))
        self.widget.resize(ln, ln2)
        self.widget.setPlainText(self.lines)
        layout = QtGui.QVBoxLayout()
        layout.addWidget(self.widget)
        layout.addLayout(buttonLayout)
        self.setLayout(layout)
        i = self.parms[0].rfind('/')
        parm = ''
        for l in range(1, len(self.parms)):
            parm += ' ' + self.parms[l]
        self.setWindowTitle('SIREN - Output from ' + self.parms[0][i + 1:])
        size = self.geometry()
        self.setGeometry(1, 1, ln + 10, ln2 + 35)
        size = self.geometry()
        self.move((screen.width() - size.width()) / 2,
            (screen.height() - size.height()) / 2)
        self.widget.show()

    def accept(self):
        try:
            i = self.parms[1].rfind('/')   # fudge to see if first parm has a directory to use as an alternative
        except:
            i = 0
        if i > 0:
            save_filename = self.parms[1][:i + 1]
        else:
            i = self.parms[0].rfind('/')
            j = self.parms[0].rfind('.')
            save_filename = self.parms[0][:j]
        for k in range(1, len(self.parms)):
            i = self.parms[k].rfind('/')
            if i > 0:
                save_filename += '_' + self.parms[k][i + 1:]
            else:
                save_filename += '_' + self.parms[k]
        save_filename += '_' + str(QtCore.QDateTime.toString(QtCore.QDateTime.currentDateTime(), 'yyyy-MM-dd_hhmm'))
        save_filename += '.txt'
        fileName = QtGui.QFileDialog.getSaveFileName(self,
                                         self.tr("QFileDialog.getSaveFileName()"),
                                         save_filename,
                                         self.tr("All Files (*);;Text Files (*.txt)"))
        if not fileName.isEmpty():
            s = open(fileName, 'w')
            s.write(self.lines)
            s.close()
        self.close()


class SAMRun():
    def __init__(self, program, parameters=None, comment=None):
    ssc_api = ssc.API()
        parms = ['python']
        parms.append(program)
        bit = None
        if parameters is not None:
            for p in parameters:
                parms.append(p)
        proc = sp.Popen(parms, stdout=sp.PIPE, stderr=sp.PIPE)
        output, err = proc.communicate()
        retcode = proc.poll()
        return_code = 'SAM SDK API:\n    SSC Version: ' + str(ssc_api.version()) + \
                      '\n    SSC Build Info: ' + ssc_api.build_info() + \
                      '\nReturn Code:\n    ' + str(retcode) + '\n'
        if retcode != 0:
            return_code += err
        dialr = RptDialog(parms[1:], return_code, output, comment=comment)
        dialr.exec_()


if "__main__" == __name__:
    app = QtGui.QApplication(sys.argv)
    fil = SAMRun('windpower_ak2.py', parameters=['Enercon E66_1870kw(MG)', '/home/' +
      getUser() +
     '/Dropbox/SEN stuff/Tech Team/SIREN App/SAM_wind_files/Albany_-35.0000_118.0000_2014.srw'])
