#!/usr/bin/python
#
#  Copyright (C) 2016 Sustainable Energy Now Inc., Angus King     
#
#  makerainfall2.py - This file is possibly part of SIREN.
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

import datetime
import gzip
from math import *
import os 
import sys
if sys.platform == 'win32' or sys.platform == 'cygwin':
    import math
    from netCDF4 import Dataset
else:
    from Scientific.N import *
    from Scientific.IO.NetCDF import *
import time
from PyQt4 import QtCore, QtGui
import displayobject
from sammodels import getDNI, getDHI
from credits import fileVersion

class makeRainfall():
    def unZip(self, inp_file):
        if inp_file[-3] == '.gz':
            if not os.path.exists(inp_file):
                self.log += time.strftime('%Y-%m-%d %H:%M:%S ',time.localtime()) + \
                           'Terminating as file not found - %s\n' % inp_file
                self.return_code = 12
                return None
            out_file = inp_file + '.nc'
            if os.path.exists(out_file):
                pass
            else:
                fin = gzip.open(inp_file, 'rb')
                file_content = fin.read()
                fin.close()
                fou = open(out_file,'wb')
                fou.write(file_content)
                fou.close()
            return out_file
        else:
            if not os.path.exists(inp_file):
                self.log += time.strftime('%Y-%m-%d %H:%M:%S ',time.localtime()) + \
                           'Terminating as file not found - %s\n' % inp_file
                self.return_code = 12
                return None
            return inp_file
    
    def getRainfall(self, prc):
        try: 
            tm = prc.data
        except:
            tm = prc[:] 
        tmp = []
        for hr in range(len(tm)): # 24 hours
            ht = []
            for lat in range(len(tm[hr])): # n latitudes
                lt = []
                for lon in range(len(tm[hr][lat])): # m longitudes
                    lt.append(round(tm[hr][lat][lon] * 3600, 2)) # second to hour
                ht.append(lt)
            tmp.append(ht)
        return tmp
    
    def valu(self, data, lat1, lon1, lat_rat, lon_rat, rnd=4):
        if rnd > 0:
            return round(lat_rat * lon_rat * data[lat1][lon1] + \
                (1.0 - lat_rat) * lon_rat * data[lat1 + 1][lon1] + \
                lat_rat * (1.0 - lon_rat) * data[lat1][lon1 + 1] + \
                (1.0 - lat_rat) * (1.0 - lon_rat) * data[lat1 + 1][lon1 + 1],rnd)
        else:
            return int(lat_rat * lon_rat * data[lat1][lon1] + \
                (1.0 - lat_rat) * lon_rat * data[lat1 + 1][lon1] + \
                lat_rat * (1.0 - lon_rat) * data[lat1][lon1 + 1] + \
                (1.0 - lat_rat) * (1.0 - lon_rat) * data[lat1 + 1][lon1 + 1])
    
    def get_data(self, inp_file):
        unzip_file = self.unZip(inp_file)
        if self.return_code != 0:
            return
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            cdf_file = Dataset(unzip_file, 'r')
        else:
            cdf_file = NetCDFFile(unzip_file, 'r')
    #   Variable Description                                          Units
    #   -------- ---------------------------------------------------- ----------
    #   prectot  Total surface precipitation flux                     kg (m)-2 (s)-1
        self.latitude = cdf_file.variables[self.vars['latitude']]
        self.lati = self.latitude[:]
        del self.latitude
        self.longitude = cdf_file.variables[self.vars['longitude']]
        self.longi = self.longitude[:]
        del self.longitude
        times = cdf_file.variables['time']
        self.tims = times[:]
        self.rain += self.getRainfall(cdf_file.variables[self.vars['prectot']])
        cdf_file.close()

    def checkZone(self):
        self.return_code = 0
        if self.longrange[0] is not None:
            if int(round(self.longrange[0] / 15)) != self.src_zone:
                self.log += 'MERRA west longitude (%s) in different time zone: %s\n' % ('{:0.4f}'.format(self.longrange[0]), \
                            int(round(self.longrange[0] / 15)))
                self.return_code = 1
        if self.longrange[1] is not None:
            if int(round(self.longrange[1] / 15)) != self.src_zone:
                self.log += 'MERRA east longitude (%s) in different time zone: %s\n' % ('{:0.4f}'.format(self.longrange[1]), \
                            int(round(self.longrange[1] / 15)))
                self.return_code = 1  
        return
 
    def close(self):
        return
    
    def getLog(self):
        return self.log    
        
    def returnCode(self):
        return str(self.return_code)

    def getInfo(self, inp_file):
        if not os.path.exists(inp_file):
            if inp_file.find('MERRA300') >= 0: 
                inp_file = inp_file.replace('MERRA300', 'MERRA301')
            else:
                return    
        if not os.path.exists(inp_file):
            return
        unzip_file = self.unZip(inp_file)
        if self.return_code != 0:
            return
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            cdf_file = Dataset(unzip_file, 'r')
        else:
            cdf_file = NetCDFFile(unzip_file, 'r')
        i = unzip_file.rfind('/')
        j = unzip_file.rfind('\\')
        if i < 0:
            i = j
        self.log += '\nFile:\n    '
        self.log += unzip_file[i + 1:] + '\n'
        self.log += ' Format:\n    '
        self.log += str(cdf_file.Format) + '\n'
        self.log += ' Dimensions:\n    '
        vals = ''
        keys = cdf_file.dimensions.keys()
        values = cdf_file.dimensions.values()
        if type(cdf_file.dimensions) is dict:
            for i in range(len(keys)):
                vals += keys[i] + ': ' + str(values[i]) + ', '
        else:
            for i in range(len(keys)):
                bits = str(values[i]).strip().split(' ')
                vals += keys[i] + ': ' + bits[-1] + ', '
        self.log += vals[:-2] + '\n'
        self.log += ' Variables:\n    '
        vals = ''
        for key in iter(sorted(cdf_file.variables.iterkeys())):
            vals += key + ', '
        self.log += vals[:-2] + '\n'
        latitude = cdf_file.variables[self.vars['latitude']][:]
        self.log += ' Latitudes:\n    '
        vals = ''
        lati = latitude[:]
        for val in lati:
            vals += '{:0.4f}'.format(val) + ', '
        self.log += vals[:-2] + '\n'
        longitude = cdf_file.variables[self.vars['longitude']][:]
        self.log += ' Longitudes:\n    '
        vals = ''
        longi = longitude[:]
        for val in longi:
            vals += '{:0.4f}'.format(val) + ', '
        self.log += vals[:-2] + '\n'
        cdf_file.close()   
        return
            
    def __init__(self, src_year, src_zone, src_dir, tgt_dir, fmat, src_lat_lon=None, info=False):
        self.log = ''
        self.return_code = 0
        self.src_year = int(src_year)
        self.src_dir = src_dir
        self.src_pfx = ''
        self.src_sfx = ''
        self.vars = {'latitude': 'lat', 'longitude': 'lon', 'prectot': 'PRECTOT'}
        if self.src_dir != '':
            self.src_dir += '/'
            fils = os.listdir(self.src_dir)
            for fil in fils:
                if fil.find('MERRA') >= 0:
                    j = fil.find('.tavg1_2d_')
                    if j > 0:
                        self.src_pfx = fil[:j + 17]
                        self.src_pfx = self.src_pfx.replace('MERRA301', 'MERRA300')
                        self.src_sfx = fil[j + 17 + 8:]
                        if fil[j + 10: j + 13] in ['flx', 'lnd', 'mld']:
                            fmat = fil[j + 10: j + 13]
                            break
            del fils
        if self.src_pfx.find('MERRA300') >= 0:
            for key in self.vars.keys():
                self.vars[key] = self.vars[key].lower()                
            self.vars['latitude'] = 'latitude'
            self.vars['longitude'] = 'longitude'
        self.tgt_dir = tgt_dir
        if self.tgt_dir != '':
            self.tgt_dir += '/'
        if info:
            inp_strt = '{0:04d}'.format(self.src_year) + '0101'
            # get variables from "rain" file
            inp_file = self.src_dir + self.src_pfx + inp_strt + self.src_sfx
            self.getInfo(inp_file)
            return
        if fmat not in ['flx', 'lnd', 'mld']:
            self.log += 'Invalid input file format specified - %s\n' % fmat
            self.return_code = 8
            return 
        if src_zone.lower() == 'auto':
            self.auto_zone = True
            inp_strt = '{0:04d}'.format(self.src_year) + '0101'
            # get longitude from "rainfall" file
            inp_file = self.src_dir + self.src_pfx + inp_strt + self.src_sfx
            if not os.path.exists(inp_file):
                if inp_file.find('MERRA300') >= 0: 
                    inp_file = inp_file.replace('MERRA300', 'MERRA301')
            unzip_file = self.unZip(inp_file)
            if self.return_code != 0:
                return
            if sys.platform == 'win32' or sys.platform == 'cygwin':
                cdf_file = Dataset(unzip_file, 'r')
            else:
                cdf_file = NetCDFFile(unzip_file, 'r')
            longitude = cdf_file.variables[self.vars['longitude']][:]
            self.src_zone = int(round(longitude[0] / 15))
            self.log += 'Time zone: %s based on MERRA (west) longitude (%s to %s)\n' % (str(self.src_zone), \
                        '{:0.4f}'.format(longitude[0]), '{:0.4f}'.format(longitude[-1]))
            cdf_file.close()
        else:
            self.src_zone = int(src_zone)
            self.auto_zone = False
        if src_lat_lon != '':
            self.src_lat = []
            self.src_lon = []
            latlon = src_lat_lon.split(',')
            try:
                for j in range(0, len(latlon), 2):
                    self.src_lat.append(float(latlon[j]))
                    self.src_lon.append(float(latlon[j + 1]))
            except:
                self.log += 'Error with Coordinates field'
                self.return_code = 4
                return
            self.log += 'Processing %s, %s\n' % (self.src_lat, self.src_lon)
        else:
            self.src_lat = None
            self.src_lon = None
        if self.src_dir[0] == self.src_dir[-1] and (self.src_dir[0] == '"' or \
           self.src_dir[0] == "'"):
            self.src_dir = self.src_dir[1:-1]
        if self.tgt_dir[0] == self.tgt_dir[-1] and (self.tgt_dir[0] == '"' or \
           self.tgt_dir[0] == "'"):
            self.tgt_dir = self.tgt_dir[1:-1]
        self.lati = []
        self.longi = []
        self.longrange = [None, None]
        self.tims = []
        self.rain = []
        dys = [31,28,31,30,31,30,31,31,30,31,30,31]
        if self.src_zone > 0:
            inp_strt = '{0:04d}'.format(self.src_year - 1) + '1231'
        elif self.src_zone <= 0:
            inp_strt = '{0:04d}'.format(self.src_year) + '0101'
        # get variables from "rainfall" files
        inp_file = self.src_dir + self.src_pfx + inp_strt + self.src_sfx
        if not os.path.exists(inp_file):
            if inp_file.find('MERRA300') >= 0: 
                inp_file = inp_file.replace('MERRA300', 'MERRA301')
        self.get_data(inp_file) # get rainfall data
        if self.return_code != 0:
            return
        if self.src_zone != 0:
        # last self.src_zone hours of (self.src_year -1) GMT = first self.src_zone of self.src_year
        # first self.src_zone hours of self.src_year GMT = last self.src_zone of (self.src_year - 1)
            self.rain = self.rain[-self.src_zone:]
        elif self.src_zone < 0:
        # first self.src_zone hours of self.src_year GMT = last self.src_zone of (self.src_year - 1)
            self.rain = self.rain[:self.src_zone]
        for mt in range(len(dys)):
            self.log += 'Processing month %s rainfall\n' % str(mt+1)
            for dy in range(1, dys[mt] + 1):
                if self.src_zone <= 0:
                    if mt == 0 and dy == 1:
                        continue
                inp_file = self.src_dir + self.src_pfx + \
                                '{0:04d}'.format(self.src_year) + '{0:02d}'.format(mt+1) + \
                                '{0:02d}'.format(dy) + self.src_sfx
                if not os.path.exists(inp_file):
                    if inp_file.find('MERRA300') >= 0: 
                        inp_file = inp_file.replace('MERRA300', 'MERRA301')
                self.get_data(inp_file) # get rainfall data
                if self.return_code != 0:
                    return
        if self.src_zone > 0:
            for i in range(self.src_zone): # delete last n hours
                del self.rain[-1]
        elif self.src_zone < 0:
            inp_strt = '{0:04d}'.format(self.src_year + 1) + '0101'
            inp_file = self.src_dir + self.src_pfx + inp_strt + self.src_sfx
            if not os.path.exists(inp_file):
                if inp_file.find('MERRA300') >= 0: 
                    inp_file = inp_file.replace('MERRA300', 'MERRA301')
            self.get_data(inp_file)
            if self.return_code != 0:
                return
            for i in range(24 + self.src_zone): # delete last n hours
                del self.rain[-1]
        target_dir = self.tgt_dir
        self.log += 'Target directory is %s\n' % target_dir
        if not os.path.exists(target_dir):
            self.log += 'mkdir %s\n' % target_dir
            os.makedirs(target_dir)
        if self.src_lat is not None: # specific location(s)
            for i in range(len(self.src_lat)):
                for lat2 in range(len(self.lati)):
                    if self.src_lat[i] <= self.lati[lat2]:
                        break
                for lon2 in range(len(self.longi)):
                    if self.src_lon[i] <= self.longi[lon2]:
                        break
                if self.longrange[0] is None:
                    self.longrange[0] = self.longi[lon2]
                else:
                    if self.longi[lon2] < self.longrange[0]:
                        self.longrange[0] = self.longi[lon2]
                if self.longrange[1] is None:
                    self.longrange[1] = self.longi[lon2]
                else:
                    if self.longi[lon2] > self.longrange[1]:
                        self.longrange[1] = self.longi[lon2]
                lat1 = lat2 - 1
                lat_rat = (self.lati[lat2] - self.src_lat[i]) / (self.lati[lat2] - self.lati[lat1])
                lon1 = lon2 - 1
                lon_rat = (self.longi[lon2] - self.src_lon[i]) / (self.longi[lon2] - self.longi[lon1])
                out_file = self.tgt_dir + 'rainfall_' + str(self.src_lat[i]) + '_' + \
                           str(self.src_lon[i]) + '_' + str(self.src_year) + '.csv'
                tf = open(out_file, 'w')
                hdr = 'Location,City,Region,Country,Latitude,Longitude,Time Zone,Elevation,Source\n' + \
                      'id,<city>,<state>,<country>,%s,%s,%s,0,MERRA\n' % (round(self.src_lat[i], 4), \
                      round(self.src_lon[i], 4), str(self.src_zone))
                tf.write(hdr)
                tf.write('Year,Month,Day,Hour,Rainfall (mm)' + '\n')
                mth = 0
                day = 1
                hour = 1
                for hr in range(len(self.rain)):
                    tf.write(str(self.src_year) + ',' + str(mth + 1) + ',' + str(day) + ',' + str(hour) + ',' + \
                             str(self.valu(self.rain[hr], lat1, lon1, lat_rat, lon_rat, rnd=1)) + '\n')
                    hour += 1
                    if hour > 24:
                        hour = 1
                        day += 1
                        if day > dys[mth]:
                            mth += 1
                            day = 1
                            hour = 1
                tf.close()
                self.log += '%s created\n' % out_file[out_file.rfind('/')+1:]
        else:
            for lat in range(len(self.rain[0])):
                for lon in range(len(self.rain[0][0])):
                    if self.longrange[0] is None:
                        self.longrange[0] = self.longi[lon]
                    else:
                        if self.longi[lon] < self.longrange[0]:
                            self.longrange[0] = self.longi[lon]
                    if self.longrange[1] is None:
                        self.longrange[1] = self.longi[lon]
                    else:
                        if self.longi[lon] > self.longrange[1]:
                            self.longrange[1] = self.longi[lon]
                    out_file = self.tgt_dir + 'rainfall_' + '{:0.4f}'.format(self.lati[lat]) + \
                               '_' + '{:0.4f}'.format(self.longi[lon]) + '_' + str(self.src_year) + '.csv'
                    tf = open(out_file, 'w')
                    hdr = 'Location,City,Region,Country,Latitude,Longitude,Time Zone,Elevation,Source\n' + \
                          'id,<city>,<state>,<country>,%s,%s,%s,0,MERRA\n' % (round(self.lati[lat], 4), \
                          round(self.longi[lon], 4), str(self.src_zone))
                    tf.write(hdr)
                    tf.write('Year,Month,Day,Hour,Rainfall (mm)' + '\n')
                    mth = 0
                    day = 1
                    hour = 1
                    for hr in range(len(self.rain)):
                        tf.write(str(self.src_year) + ',' + str(mth + 1) + ',' + str(day) + ',' + str(hour) + ',' + \
                                 str(self.rain[hr][lat][lon]) + '\n')
                        hour += 1
                        if hour > 24:
                            hour = 1
                            day += 1
                            if day > dys[mth]:
                                mth += 1
                                day = 1
                                hour = 1
                    tf.close()
                    self.log += '%s created\n' % out_file[out_file.rfind('/')+1:]
                    self.checkZone() 
        return # that's it  

class ClickableQLabel(QtGui.QLabel): 
    def __init(self, parent):
        QLabel.__init__(self, parent)
 
    def mousePressEvent(self, event):
        QtGui.QApplication.widgetAt(event.globalPos()).setFocus()
        self.emit(QtCore.SIGNAL('clicked()'))
         
class getParms(QtGui.QWidget):
 
    def __init__(self, help='help.html'):
        super(getParms, self).__init__()
        self.help = help
        self.initUI()
        
    def initUI(self):
        self.yearSpin = QtGui.QSpinBox()
        now = datetime.datetime.now()
        self.yearSpin.setRange(1979, now.year)
        self.yearSpin.setValue(now.year - 1)
        self.zoneCombo = QtGui.QComboBox()
        self.zoneCombo.addItem('Auto')
        for i in range(-12, 13):
            self.zoneCombo.addItem(str(i))
        self.zoneCombo.currentIndexChanged[str].connect(self.zoneChanged)
        self.zone_lon = QtGui.QLabel(('Time zone calculated from MERRA data'))
        self.grid = QtGui.QGridLayout()
        self.grid.addWidget(QtGui.QLabel('Year:'), 0, 0)
        self.grid.addWidget(self.yearSpin, 0, 1)
        self.grid.addWidget(QtGui.QLabel('Time Zone:'), 1, 0)
        self.grid.addWidget(self.zoneCombo, 1, 1)
        self.grid.addWidget(self.zone_lon, 1, 2, 1, 3)
        self.grid.addWidget(QtGui.QLabel('Rainfall File Format:'), 2, 0)
        self.fmatcombo = QtGui.QComboBox(self)
        self.fmats = ['any', 'flx', 'lnd', 'mld']
        for i in range(len(self.fmats)):
            self.fmatcombo.addItem(self.fmats[i])
        self.fmatcombo.setCurrentIndex(0)
        self.grid.addWidget(self.fmatcombo, 2, 1)
        self.grid.addWidget(QtGui.QLabel('Coordinates:'), 3, 0)
        self.coords = QtGui.QPlainTextEdit()
        self.grid.addWidget(self.coords, 3, 1, 1, 3)
        self.grid.addWidget(QtGui.QLabel('Copy folder down:'), 4, 0)
        self.checkbox = QtGui.QCheckBox()
        self.checkbox.setCheckState(QtCore.Qt.Checked)
        self.grid.addWidget(self.checkbox, 4, 1)
        self.grid.addWidget(QtGui.QLabel('If checked will copy rainfall folder changes down to others'), 4, 3)
        cur_dir = os.getcwd()
        self.dir_labels = ['Rainfall Source', 'Target']
        self.dirs = [None, None, None]
        for i in range(2):
            self.grid.addWidget(QtGui.QLabel(self.dir_labels[i] + ' Folder:'), 5 + i, 0)
            self.dirs[i] = ClickableQLabel()
            self.dirs[i].setText(cur_dir)
            self.dirs[i].setFrameStyle(6)
            self.connect(self.dirs[i], QtCore.SIGNAL('clicked()'), self.dirChanged)
            self.grid.addWidget(self.dirs[i], 5 + i, 1, 1, 3)
        quit = QtGui.QPushButton('Quit', self)
        self.grid.addWidget(quit, 8, 0)
        quit.clicked.connect(self.quitClicked) 
        QtGui.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
        dorain = QtGui.QPushButton('Produce Rainfall Files', self)
        wdth = dorain.fontMetrics().boundingRect(dorain.text()).width() + 9
        self.grid.addWidget(dorain, 8, 1)
        dorain.clicked.connect(self.dorainClicked)
        help = QtGui.QPushButton('Help', self)
        help.setMaximumWidth(wdth)
        self.grid.addWidget(help, 8, 2)
        help.clicked.connect(self.helpClicked)
        QtGui.QShortcut(QtGui.QKeySequence('F1'), self, self.helpClicked)
        info = QtGui.QPushButton('File Info', self)
        info.setMaximumWidth(wdth)
        self.grid.addWidget(info, 8, 3)
        info.clicked.connect(self.infoClicked)
     #   self.grid.setColumnStretch(4, 2)
        frame = QtGui.QFrame()
        frame.setLayout(self.grid)
        self.scroll = QtGui.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtGui.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
        self.setWindowTitle('SIREN makerainfall2 (' + fileVersion() + ') - Make rainfall files from MERRA data')
        self.center()
        self.show()

    def center(self):
        frameGm = self.frameGeometry()
        screen = QtGui.QApplication.desktop().screenNumber(QtGui.QApplication.desktop().cursor().pos())
        centerPoint = QtGui.QApplication.desktop().availableGeometry(screen).center()
        frameGm.moveCenter(centerPoint)
        self.move(frameGm.topLeft())

    def zoneChanged(self, val):
        if self.zoneCombo.currentIndex() == 0:
            lon = '(Time zone calculated from MERRA data)'
        else:
            lw = int(self.zoneCombo.currentIndex() - 13) * 15 - 7.5
            le = int(self.zoneCombo.currentIndex() - 13) * 15 + 7.5
            if lw < -180:
                lw = lw + 360
            if le > 180:
                le = le - 360
            lon = '(Approx. Longitude: %s to %s)' % ('{:0.1f}'.format(lw), '{:0.1f}'.format(le))
        self.zone_lon.setText(lon)

    def zoomChanged(self, val):
        self.zoomScale.setText('(' + scale[int(val)] + ')')
        self.zoomScale.adjustSize()

    def dirChanged(self):
        for i in range(3):
            if self.dirs[i].hasFocus():
                break
        curdir = self.dirs[i].text()
        newdir = str(QtGui.QFileDialog.getExistingDirectory(self, 'Choose ' + self.dir_labels[i] + ' Folder',
                 curdir, QtGui.QFileDialog.ShowDirsOnly))
        if newdir != '':
            self.dirs[i].setText(newdir) 
            if self.checkbox.isChecked():
                if i == 0:
                    self.dirs[1].setText(newdir) 

    def helpClicked(self):   
        dialog = displayobject.AnObject(QtGui.QDialog(), self.help, \
                 title='Help for SIREN makerainfall2 (' + fileVersion() + ')', section='merrar')
        dialog.exec_()

    def quitClicked(self):      
        self.close()

    def dorainClicked(self):
        coords = str(self.coords.toPlainText())
        if coords != '':
            if coords.find(' ') >= 0:
                if coords.find(',') < 0:
                    coords = coords.replace(' ', ',')
                else:
                    coords = coords.replace(' ', '')
        if self.zoneCombo.currentIndex() == 0:
            zone = 'auto'
        else:
            zone = str(self.zoneCombo.currentIndex() - 13)
        rain = makeRainfall(str(self.yearSpin.value()), zone, str(self.dirs[0].text()), \
                            str(self.dirs[1].text()), str(self.fmatcombo.currentText()), coords)
        dialr = RptDialog(str(self.yearSpin.value()), zone, str(self.dirs[0].text()), \
                          str(self.dirs[1].text()), str(self.fmatcombo.currentText()), \
                          coords, rain.returnCode(), rain.getLog())
        dialr.exec_()
        del rain
        del dialr
        
    def infoClicked(self):
        coords = str(self.coords.toPlainText())
        if coords != '':
            if coords.find(' ') >= 0:
                if coords.find(',') < 0:
                    coords = coords.replace(' ', ',')
                else:
                    coords = coords.replace(' ', '')
        if self.zoneCombo.currentIndex() == 0:
            zone = 'auto'
        else:
            zone = str(self.zoneCombo.currentIndex() - 13)
        rain = makeRainfall(str(self.yearSpin.value()), zone, str(self.dirs[0].text()), \
                            str(self.dirs[1].text()), str(self.fmatcombo.currentText()), coords, info=True)
        dialr = RptDialog(str(self.yearSpin.value()), zone, str(self.dirs[0].text()), \
                          str(self.dirs[1].text()), str(self.fmatcombo.currentText()), \
                          coords, rain.returnCode(), rain.getLog())
        dialr.exec_()
        del rain
        del dialr

        
class RptDialog(QtGui.QDialog):
    def __init__(self, year, zone, src_dir, tgt_dir, fmat, coords, return_code, output):
        super(RptDialog, self).__init__()
        self.parms = [str(year), str(zone), fmat, tgt_dir, src_dir]
        max_line = 0
        self.lines = 'Parameters:\n    Year: %s\n    Time Zone: %s\n    Input Format: %s\n' % \
                     (year, zone, fmat) 
        self.lines += '    Rainfall Files: %s\n' % src_dir
        self.lines += '    Target Folder: %s\n' % tgt_dir
        if coords != '':
            self.lines += '    Coordinates: %s\n' % coords
        self.lines += 'Return Code:\n    %s\n' % return_code
        self.lines += 'Output:\n'
        self.lines += output
        lenem = self.lines.split('\n')
        line_cnt = len(lenem)
        for i in range(line_cnt):
            max_line = max(max_line, len(lenem[i]))
        del lenem
        QtGui.QDialog.__init__(self)
        self.saveButton = QtGui.QPushButton(self.tr('&Save'))
        self.cancelButton = QtGui.QPushButton(self.tr('Cancel'))
        buttonLayout = QtGui.QHBoxLayout()
        buttonLayout.addStretch(1)
        buttonLayout.addWidget(self.saveButton)
        buttonLayout.addWidget(self.cancelButton)
        self.connect(self.saveButton, QtCore.SIGNAL('clicked()'), self, \
                     QtCore.SLOT('accept()'))
        self.connect(self.cancelButton, QtCore.SIGNAL('clicked()'), \
                     self, QtCore.SLOT('reject()'))
        self.widget = QtGui.QTextEdit()
        self.widget.setFont(QtGui.QFont('Courier New', 11))
        fnt = self.widget.fontMetrics()
        ln = (max_line + 5) * fnt.maxWidth()
        ln2 = (line_cnt + 2) * fnt.height() 
        screen = QtGui.QDesktopWidget().availableGeometry()
        if ln > screen.width() * .67:
            ln = int(screen.width() * .67)
        if ln2 > screen.height() * .67:
            ln2 = int(screen.height() * .67)
        self.widget.setSizePolicy(QtGui.QSizePolicy(QtGui.QSizePolicy.Expanding, \
            QtGui.QSizePolicy.Expanding)) 
        self.widget.resize(ln, ln2)
        self.widget.setPlainText(self.lines)
        layout = QtGui.QVBoxLayout()
        layout.addWidget(self.widget)
        layout.addLayout(buttonLayout)
        self.setLayout(layout)
        self.setWindowTitle('Output from makerainfall2')
        size = self.geometry()
        self.setGeometry(1, 1 , ln + 10, ln2 + 35)
        size = self.geometry()
        self.move((screen.width()-size.width())/2, \
            (screen.height()-size.height())/2)
        self.widget.show()

    def accept(self):
        i = sys.argv[0].rfind('/') # fudge to see if program has a directory to use as an alternative
        j = sys.argv[0].rfind('.')
        if i > 0:
            save_filename = sys.argv[0][i+1:j]
        else:
            save_filename = sys.argv[0][:j]
        last_bit = ''
        for k in range(len(self.parms)):
            if self.parms[k] == '':
                continue
            i = self.parms[k].rfind('/')
            if i > 0:
                if self.parms[k][i+1:] != last_bit:
                    save_filename += '_' + self.parms[k][i+1:]
                    last_bit = self.parms[k][i+1:]
            else:
                if self.parms[k] != last_bit:
                    save_filename += '_' + self.parms[k]
                    last_bit = self.parms[k]
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


if "__main__" == __name__:
    app = QtGui.QApplication(sys.argv)
    if len(sys.argv) > 1: # arguments
        src_lat_lon = ''
        src_year = 2014
        src_zone = 0
        fmat = 'lnd'
        src_dir = ''
        tgt_dir = ''
        for i in range(1, len(sys.argv)):
            if sys.argv[i][:5] == 'year=':
                src_year = int(sys.argv[i][5:])
            elif sys.argv[i][:7] == 'latlon=' or sys.argv[i][:7] == 'coords=': # lat and lon
                src_lat_lon = sys.argv[i][7:]
            elif sys.argv[i][:5] == 'zone=':
                src_zone = int(sys.argv[i][5:])
            elif sys.argv[i][:9] == 'timezone=':
                src_zone = int(sys.argv[i][9:])
            elif sys.argv[i][:5] == 'fmat=':
                fmat = sys.argv[i][5:]
            elif sys.argv[i][:7] == 'format=':
                fmat = sys.argv[i][7:]
            elif sys.argv[i][:7] == 'source=' or sys.argv[i][:7] == 'srcdir=':
                src_dir = sys.argv[i][7:]
            elif sys.argv[i][:7] == 'target=' or sys.argv[i][:7] == 'tgtdir=':
                tgt_dir = sys.argv[i][7:]
        files = makeRainfall(src_year, src_zone, src_dir, tgt_dir, src_lat_lon)
        dialr = RptDialog(files.returnCode(), files.getLog())
        dialr.exec_()
    else:
        ex = getParms()
        app.exec_()
        app.deleteLater()
        sys.exit()
