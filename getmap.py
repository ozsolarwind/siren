#!/usr/bin/python3
#
#  Copyright (C) 2015-2025 Sustainable Energy Now Inc., Angus King
#
#  getmap.py - This file is part of SIREN.
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

from urllib3 import PoolManager
import math
import os
import sys
import tempfile
import configparser   # decode .ini file
from PyQt5 import QtCore, QtGui, QtWidgets
import displayobject
from credits import fileVersion
from getmodels import getModelFile
from senutils import ClickableQLabel
import worldwindow

scale = {0: '1:500 million', 1: '1:250 million', 2: '1:150 million', 3: '1:70 million',
         4: '1:35 million', 5: '1:15 million', 6: '1:10 million', 7: '1:4 million',
         8: '1:2 million', 9: '1:1 million', 10: '1:500,000', 11: '1:250,000',
         12: '1:150,000', 13: '1:70,000', 14: '1:35,000', 15: '1:15,000', 16: '1:8,000',
         17: '1:4,000', 18: '1:2,000', 19: '1:1,000'}


class retrieveMap():

    def mercatorToLat(self, mercatorY):
        return(math.degrees(math.atan(math.sinh(mercatorY))))

    def latEdges(self, y, z):
        n = 2 ** z
        unit = 1. / n
        relY1 = y * unit
        relY2 = relY1 + unit
        if (1 - 2 * relY1) == 1.:
            lat1 = 90.
        else:
            lat1 = self.mercatorToLat(math.pi * (1 - 2 * relY1))
        if (1 - 2 * relY2) == -1.:
            lat2 = -90.
        else:
            lat2 = self.mercatorToLat(math.pi * (1 - 2 * relY2))
        return(lat1, lat2)

    def lonEdges(self, x, z):
        n = 2 ** z
        unit = 360. / n
        lon1 = -180. + x * unit
        lon2 = lon1 + unit
        return(lon1, lon2)

    def tileEdges(self, x, y, z):
        lat1, lat2 = self.latEdges(y, z)
        lon1, lon2 = self.lonEdges(x, z)
        return((lat2, lon1, lat1, lon2))   # S,W,N,E

    def deg2num(self, lat_deg, lon_deg, zoom):
        # Derived from:  https://wiki.openstreetmap.org/wiki/Slippy_map_tilenames#Python
        if lat_deg > 85:
            lat_deg = 85
        elif lat_deg < -85:
            lat_deg = -85
        lat_rad = math.radians(lat_deg)
        n = 2.0 ** zoom
        xtile = int((lon_deg + 180.0) / 360.0 * n)
        if xtile >= n:
            xtile = int(n - 1)
        try:
            ytile = int((1.0 - math.asinh(math.tan(lat_rad)) / math.pi) / 2.0 * n)
        except:
            ytile = 0
        return (xtile, ytile)

    def writetile(self, x, y, zoom):
        url = self.url
        if len(self.subs) > 0:
            self.sub_ctr += 1
            if self.sub_ctr >= len(self.subs):
                self.sub_ctr = 0
            url = url.replace('[]', self.subs[self.sub_ctr])
        file_name = str(x) + '_' + str(y) + '.png'
        url_tail = self.url_tail.replace('/zoom', '/' + str(zoom))
        url_tail = url_tail.replace('/x', '/' + str(x))
        url_tail = url_tail.replace('/y', '/' + str(y))
        if 1 == 2: # Google?
            url_tail = self.url_tail.replace('z=zoom', 'z=' + str(zoom))
            url_tail = url_tail.replace('x=x', 'x=' + str(x))
            url_tail = url_tail.replace('y=y', 'y=' + str(y))
        if self.batch and self.debug:
            print(url + url_tail)
        user_agent = {'User-agent': 'getmap ' + fileVersion() + ' contact siren@sen.asn.au'}
        http = PoolManager(headers=user_agent)
        if self.batch and self.debug:
            print('Retrieving ' + url_tail)
        response = http.request('GET', url + url_tail)
        if response.status == 200 and response.reason == 'OK':
            if self.batch and self.debug:
                print(url_tail + ' retrieved')
            f = open(self.tmp_location + file_name, 'wb')
            f.write(response.data)
            f.close()
            message = 'OK'
        else:
            message = url_tail + ' failed\n' + str(response.status) + ' ' + response.reason
        return message, file_name

    def __init__(self, upper_lat, upper_lon, lower_lat, lower_lon, output, zoom=None, url=None, width=None, height=None, caller=None, debug=None):
        if len(sys.argv) > 1 and sys.argv[1][-4:] != '.ini':
            self.batch = True
        else:
            self.batch = False
            self.caller = caller
        self.debug = debug
        self.log = ''
        self.properties = ''
        config_file = getModelFile('getfiles.ini')
        config = configparser.RawConfigParser()
        config.read(config_file)
        if width != None and height != None: # Mapquest map
            try:
                url = config.get('getmap', 'mapquest_url')
            except:
                url = 'www.mapquestapi.com'
            try:
                tail = config.get('getmap', 'mapquest_tail')
            except:
                tail = '/staticmap/v4/getmap?type=sat&margin=0&bestfit=%s,%s,%s,%s&size=%s,%s&imagetype=%s'
            url_tail = tail % (upper_lat, upper_lon, lower_lat, lower_lon,
                       width, height, output[output.rfind('.') + 1:])
            try:
                url_key = '&key=' + config.get('getmap', 'mapquest_key')
            except:
                url_key = '&key=yWspjYHSK6FHtNLzZVcqP3WBxSWSwEo8'
            if self.batch and self.debug:
                print(url + url_tail)
            user_agent = {'User-agent': 'getmap ' + fileVersion() + ' contact siren@sen.asn.au'}
            http = PoolManager(headers=user_agent)
            if self.batch:
                print('Requesting ' + url_tail)
            response = http.request('GET', url + url_tail + url_key)
            if response.status == 200 and response.reason == 'OK':
                if self.batch:
                    print(url_tail + ' retrieved')
                f = open(output, 'wb')
                f.write(response.data)
                f.close()
                self.log += '\nSaving map to ' + output
                self.properties = 'map_choice=?'
                self.properties += '\nmap=%s' % (output)
                self.properties += '\nupper_left=%1.3f, %1.3f' % (upper_lat, upper_lon)
                self.properties += '\nlower_right=%1.3f, %1.3f' % (lower_lat, lower_lon)
            else:
                if self.batch:
                    print(url_tail + ' failed')
                    print(str(response.status) + ' ' + response.reason)
            return
        if url is None:
            try:
                self.url = config.get('getmap', 'url_template')
            except:
                self.url = 'https://[abc].tile.openstreetmap.org/zoom/x/y.png'
        else:
            self.url = url
        top_left = self.deg2num(upper_lat, upper_lon, zoom)
        bottom_right = self.deg2num(lower_lat, lower_lon, zoom)
        height = (bottom_right[1] - top_left[1] + 1) * 256
        width = (bottom_right[0] - top_left[0] + 1) * 256
        st, wt, nt, et = self.tileEdges(top_left[0], top_left[1], zoom)
        sb, wb, nb, eb = self.tileEdges(bottom_right[0], bottom_right[1], zoom)
        if self.batch and self.debug:
            print('(188)', '%d: %d,%d --> %1.3f :: %1.3f, %1.3f :: %1.3f' % (zoom, top_left[0], top_left[1], st, nt, wt, et))
            print('(189)', '%d: %d,%d --> %1.3f :: %1.3f, %1.3f :: %1.3f' % (zoom, bottom_right[0], bottom_right[1], sb, nb, wb, eb))
        w = bottom_right[0] - top_left[0] + 1
        h = bottom_right[1] - top_left[1] + 1
        if self.batch:
            print(w, 'x', h, '=', w * h, 'tiles.', w * 256, 'x', h * 256, 'pixels (approx.', w * 256 * h * 256, 'uncompressed bytes)')
            print('map_choice=%s' % (zoom))
            print('map%s=%s' % (zoom, output))
            print('upper_left%d=%1.3f, %1.3f' % (zoom, nt, wt))
            print('lower_right%d=%1.3f, %1.3f' % (zoom, sb, eb))
            if output == '?' or output == '':
                sys.exit()
        else:
            self.log = '%s x %s = %s tiles. %s x %s pixels (approx. %s uncompressed bytes)' % (w, h, w * h,
                       '{:,}'.format(w * 256), '{:,}'.format(h * 256), "{:,}".format(w * 256 * h * 256))
            self.properties = 'map_choice=%s' % (zoom)
            self.properties += '\nmap%s=%s' % (zoom, output)
            self.nt, self.wt, self.sb, self.eb = nt, wt, sb, eb
            self.properties += '\nupper_left%d=%1.3f, %1.3f' % (zoom, nt, wt)
            self.properties += '\nlower_right%d=%1.3f, %1.3f' % (zoom, sb, eb)
            if output == '?' or output == '':
                return
        i = self.url.find('//')
        if i >= 0:
            self.url = self.url[i + 2:]
        i = self.url.find('/')
        self.url_tail = self.url[i:]
        self.url = self.url[:i]
        self.subs = []
        i = self.url.find('[')
        if i >= 0:
            j = self.url.find(']', i)
            if j > 0:
                for k in range(i + 1, j):
                    self.subs.append(self.url[k])
                if i > 0:
                    self.url = self.url[:i + 1] + self.url[j:]
                else:
                    self.url = self.url[0] + self.url[j:]
        self.sub_ctr = -1
        self.tmp_location = tempfile.gettempdir() + '/'
        i = output.rfind('.')
        if i < 0:
            fname = output + '.png'
        else:
            fname = output
        outputimg = QtGui.QPixmap(width, height)
        painter = QtGui.QPainter(outputimg)
        if self.batch:
            print('Saving map to ' + fname)
        else:
            self.log += '\nSaving map to ' + fname
            tl = 0
            self.caller.progressbar.setMaximum((bottom_right[0] - top_left[0] + 1) * (bottom_right[1] - top_left[1] + 1) - 1)
            self.caller.progresslabel.setText('Downloading tiles')
        for x in range(top_left[0], bottom_right[0] + 1):
            for y in range(top_left[1], bottom_right[1] + 1):
                msg, tile = self.writetile(x, y, zoom)
                if msg != 'OK':
                    if self.batch:
                        print(msg)
                    else:
                        self.log += '\n' + msg
                    painter.end()
                    return
                foo = QtGui.QImage(self.tmp_location + tile)
                painter.drawImage(QtCore.QPoint(256 * (x - top_left[0]), 256 * (y - top_left[1])), foo)
                if not self.batch:
                    tl += 1
                    self.caller.progressbar.setValue(tl)
        outputimg.save(fname, fname[i + 1:])
        painter.end()
        if len(sys.argv) == 1:
            self.log += '\nDone'

    def getLog(self):
        return self.log

    def getProperties(self):
        return self.properties

    def getCoords(self):
        return [self.nt, self.wt, self.sb, self.eb]


class getMap(QtWidgets.QWidget):
    statusmsg = QtCore.pyqtSignal()

    def __init__(self, help='help.html'):
        super(getMap, self).__init__()
        self.help = help
        self.ignore = False
        self.worldwindow = None
        self.northSpin = QtWidgets.QDoubleSpinBox()
        self.northSpin.setDecimals(3)
        self.northSpin.setSingleStep(.5)
        self.northSpin.setRange(-90, 90)
        self.westSpin = QtWidgets.QDoubleSpinBox()
        self.westSpin.setDecimals(3)
        self.westSpin.setSingleStep(.5)
        self.westSpin.setRange(-180, 180)
        self.southSpin = QtWidgets.QDoubleSpinBox()
        self.southSpin.setDecimals(3)
        self.southSpin.setSingleStep(.5)
        self.southSpin.setRange(-90, 90)
        self.eastSpin = QtWidgets.QDoubleSpinBox()
        self.eastSpin.setDecimals(3)
        self.eastSpin.setSingleStep(.5)
        self.eastSpin.setRange(-180, 180)
        if len(sys.argv) > 1:
            his_config_file = getModelFile(sys.argv[1])
            his_config = configparser.RawConfigParser()
            his_config.read(his_config_file)
            try:
                mapp = his_config.get('Map', 'map_choice')
            except:
                mapp = ''
            try:
                 upper_left = his_config.get('Map', 'upper_left' + mapp).split(',')
                 self.northSpin.setValue(float(upper_left[0].strip()))
                 self.westSpin.setValue(float(upper_left[1].strip()))
                 lower_right = his_config.get('Map', 'lower_right' + mapp).split(',')
                 self.southSpin.setValue(float(lower_right[0].strip()))
                 self.eastSpin.setValue(float(lower_right[1].strip()))
            except:
                 try:
                     lower_left = his_config.get('Map', 'lower_left' + mapp).split(',')
                     upper_right = his_config.get('Map', 'upper_right' + mapp).split(',')
                     self.northSpin.setValue(float(upper_right[0].strip()))
                     self.westSpin.setValue(float(lower_left[1].strip()))
                     self.southSpin.setValue(float(lower_left[0].strip()))
                     self.eastSpin.setValue(float(upper_right[1].strip()))
                 except:
                     pass
        self.northSpin.valueChanged.connect(self.showArea)
        self.westSpin.valueChanged.connect(self.showArea)
        self.southSpin.valueChanged.connect(self.showArea)
        self.eastSpin.valueChanged.connect(self.showArea)
        self.grid = QtWidgets.QGridLayout()
        self.grid.addWidget(QtWidgets.QLabel('Area of Interest:'), 0, 0)
        area = QtWidgets.QPushButton('Choose area via Map', self)
        self.grid.addWidget(area, 0, 1, 1, 2)
        area.clicked.connect(self.areaClicked)
        self.grid.addWidget(QtWidgets.QLabel('Upper left:'), 1, 0)
        self.grid.itemAt(self.grid.count() - 1).setAlignment(QtCore.Qt.AlignRight)
        self.grid.addWidget(QtWidgets.QLabel('  North'), 2, 0)
        self.grid.itemAt(self.grid.count() - 1).setAlignment(QtCore.Qt.AlignRight)
        self.grid.addWidget(self.northSpin, 2, 1)
        self.grid.addWidget(QtWidgets.QLabel('  West'), 3, 0)
        self.grid.itemAt(self.grid.count() - 1).setAlignment(QtCore.Qt.AlignRight)
        self.grid.addWidget(self.westSpin, 3, 1)
        self.grid.addWidget(QtWidgets.QLabel('Lower right:'), 1, 2)
        self.grid.itemAt(self.grid.count() - 1).setAlignment(QtCore.Qt.AlignRight)
        self.grid.addWidget(QtWidgets.QLabel('  South'), 2, 2)
        self.grid.itemAt(self.grid.count() - 1).setAlignment(QtCore.Qt.AlignRight)
        self.grid.addWidget(self.southSpin, 2, 3)
        self.grid.addWidget(QtWidgets.QLabel('  East'), 3, 2)
        self.grid.itemAt(self.grid.count() - 1).setAlignment(QtCore.Qt.AlignRight)
        self.grid.addWidget(self.eastSpin, 3, 3)
        self.grid.addWidget(QtWidgets.QLabel('Approx. area:'), 4, 0)
        self.approx_area = QtWidgets.QLabel('')
        self.grid.addWidget(self.approx_area, 4, 1, 1, 2)
        zoom = QtWidgets.QLabel('Map Scale (Zoom):')
        self.zoomSpin = QtWidgets.QSpinBox()
        self.zoomSpin.setValue(6)
        config_file = getModelFile('getfiles.ini')
        config = configparser.RawConfigParser()
        config.read(config_file)
        try:
            maxz = int(config.get('getmap', 'max_zoom'))
        except:
            maxz = 11
        self.zoomSpin.setRange(0, maxz)
        self.zoomSpin.valueChanged[str].connect(self.zoomChanged)
        self.zoomScale = QtWidgets.QLabel('(' + scale[6] + ')')
        self.grid.addWidget(zoom, 5, 0)
        self.grid.addWidget(self.zoomSpin, 5, 1)
        self.grid.addWidget(self.zoomScale, 5, 2, 1, 2)
        self.grid.addWidget(QtWidgets.QLabel('URL template:'), 6, 0)
        self.urltemplate = QtWidgets.QLineEdit()
        try:
            url = his_config.get('getmap', 'url_template')
        except:
            url = 'https://[abc].tile.openstreetmap.org/zoom/x/y.png'
        self.urltemplate.setText(url)
        self.grid.addWidget(self.urltemplate, 6, 1, 1, 5)
        self.grid.addWidget(QtWidgets.QLabel('Image Width:'), 7, 0)
        self.widthSpin = QtWidgets.QSpinBox()
        self.widthSpin.setSingleStep(50)
        self.widthSpin.setRange(50, 3840)
        self.widthSpin.setValue(400)
        self.grid.addWidget(self.widthSpin, 7, 1)
        hlabel = QtWidgets.QLabel('Image Height:   ')
        hlabel.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
        self.grid.addWidget(hlabel, 7, 2)
        self.heightSpin = QtWidgets.QSpinBox()
        self.heightSpin.setSingleStep(50)
        self.heightSpin.setRange(50, 3840)
        self.heightSpin.setValue(200)
        self.grid.addWidget(self.heightSpin, 7, 3)
        adate = QtWidgets.QDateEdit()
        dw = adate.minimumSizeHint().width()
        sw = self.northSpin.minimumSizeHint().width()
        if sw > dw: # fix for wide QDoubleSpinBox width in Windows
            self.northSpin.setMinimumWidth(adate.minimumSizeHint().width())
            self.westSpin.setMinimumWidth(adate.minimumSizeHint().width())
            self.southSpin.setMinimumWidth(adate.minimumSizeHint().width())
            self.eastSpin.setMinimumWidth(adate.minimumSizeHint().width())
        sw = self.heightSpin.minimumSizeHint().width()
        if sw > dw: # fix for wide QSpinBox width in Windows
            self.zoomSpin.setMinimumWidth(adate.minimumSizeHint().width())
            self.widthSpin.setMinimumWidth(adate.minimumSizeHint().width())
            self.heightSpin.setMinimumWidth(adate.minimumSizeHint().width())
        del adate
        self.grid.addWidget(QtWidgets.QLabel('Image File name:'), 8, 0)
        cur_dir = os.getcwd()
        self.filename = ClickableQLabel()
        self.filename.setStyleSheet("background-color: white; border: 1px inset grey; min-height: 22px; border-radius: 4px;")
        self.filename.setText(cur_dir + '/untitled.png')
        self.filename.clicked.connect(self.fileChanged)
        self.grid.addWidget(self.filename, 8, 1, 1, 5)
        self.grid.addWidget(QtWidgets.QLabel('Properties:'), 9, 0)
        self.properties = QtWidgets.QPlainTextEdit()
        self.properties.setMaximumHeight(int(self.northSpin.sizeHint().height() * 4.5))
        self.properties.setReadOnly(True)
        self.grid.addWidget(self.properties, 9, 1, 3, 5)
        self.log = QtWidgets.QLabel()
        self.grid.addWidget(self.log, 14, 1, 3, 5)
        self.progressbar = QtWidgets.QProgressBar()
        self.progressbar.setMinimum(0)
        self.progressbar.setMaximum(100)
        self.progressbar.setValue(0)
        self.progressbar.setStyleSheet('QProgressBar {border: 1px solid grey; border-radius: 2px; text-align: center;}' \
                                       + 'QProgressBar::chunk { background-color: #06A9D6;}')
        self.grid.addWidget(self.progressbar, 17, 1, 1, 5)
        self.progressbar.setHidden(True)
        self.progresslabel=QtWidgets.QLabel('')
        self.grid.addWidget(self.progresslabel, 17, 1, 1, 4)
        self.progresslabel.setHidden(True)
        quit = QtWidgets.QPushButton('Quit', self)
        self.grid.addWidget(quit, 18, 0)
        quit.clicked.connect(self.quitClicked)
        QtWidgets.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
        query = QtWidgets.QPushButton('Query Map', self)
        wdth = query.fontMetrics().boundingRect(query.text()).width() + 9
        self.grid.addWidget(query, 18, 1)
        query.clicked.connect(self.queryClicked)
        make = QtWidgets.QPushButton('Get Map', self)
        make.setMaximumWidth(wdth)
        self.grid.addWidget(make, 18, 2)
        make.clicked.connect(self.makeClicked)
        mapquest = QtWidgets.QPushButton('MapQuest', self)
        mapquest.setMaximumWidth(wdth)
        self.grid.addWidget(mapquest, 18, 3)
        mapquest.clicked.connect(self.mapquestClicked)
        help = QtWidgets.QPushButton('Help', self)
        help.setMaximumWidth(wdth)
        self.grid.addWidget(help, 18, 4)
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
        self.setWindowTitle('SIREN - getmap (' + fileVersion() + ") - Make Map from OSM or MapQuest")
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

    def fileChanged(self):
        self.filename.setText(QtWidgets.QFileDialog.getSaveFileName(self, 'Save Image File',
                              self.filename.text(), 'Images (*.jpeg *.jpg *.png)')[0])
        if self.filename.text() != '':
            i = self.filename.text().rfind('.')
            if i < 0:
                self.filename.setText(self.filename.text() + '.png')

    def helpClicked(self):
        dialog = displayobject.AnObject(QtWidgets.QDialog(), self.help,
                 title='Help for getmap (' + fileVersion() + ')', section='map')
        dialog.exec_()

    def quitClicked(self):
        self.close()

    def queryClicked(self):
        if self.northSpin.value() < self.southSpin.value():
            l = self.northSpin.value()
            self.northSpin.setValue(self.southSpin.value())
            self.southSpin.setValue(l)
        if self.eastSpin.value() < self.westSpin.value():
            l = self.eastSpin.value()
            self.eastSpin.setValue(self.westSpin.value())
            self.westSpin.setValue(l)
        mapp = retrieveMap(self.northSpin.value(), self.westSpin.value(), self.southSpin.value(), self.eastSpin.value(),
                           '?', self.zoomSpin.value(), url=self.urltemplate.text(), caller=self)
        self.properties.setPlainText(mapp.getProperties())
        self.log.setText(mapp.getLog())

    def makeClicked(self):
        self.progressbar.setHidden(False)
        self.progresslabel.setText('')
        self.progresslabel.setHidden(False)
        mapp = retrieveMap(self.northSpin.value(), self.westSpin.value(), self.southSpin.value(), self.eastSpin.value(),
                           self.filename.text(), zoom=self.zoomSpin.value(), url=self.urltemplate.text(), caller=self)
        self.progressbar.setValue(0)
        self.progressbar.setHidden(True)
        self.progresslabel.setHidden(True)
        self.properties.setPlainText(mapp.getProperties())
        self.log.setText(mapp.getLog())

    def mapquestClicked(self):
        self.progressbar.setHidden(False)
        self.progresslabel.setText('')
        self.progresslabel.setHidden(False)
        mapp = retrieveMap(self.northSpin.value(), self.westSpin.value(), self.southSpin.value(), self.eastSpin.value(),
               self.filename.text(), width=self.widthSpin.value(), height=self.heightSpin.value(), caller=self)
        self.progressbar.setValue(0)
        self.progressbar.setHidden(True)
        self.progresslabel.setHidden(True)
        self.properties.setPlainText(mapp.getProperties())
        self.log.setText(mapp.getLog())

    @QtCore.pyqtSlot(list, str)
    def maparea(self, rectangle, approx_area=None):
        if type(rectangle) is str:
            if rectangle == 'goodbye':
                 self.worldwindow = None
                 return
        elif type(rectangle[0]) is str:
            if rectangle[0] == 'goodbye':
                 self.worldwindow = None
                 return
        self.ignore = True
        self.northSpin.setValue(rectangle[0].y())
        self.westSpin.setValue(rectangle[0].x())
        self.southSpin.setValue(rectangle[1].y())
        self.eastSpin.setValue(rectangle[1].x())
        self.ignore = False
        self.approx_area.setText(approx_area)

    def areaClicked(self):
        if self.worldwindow is None:
            scene = worldwindow.WorldScene()
            self.worldwindow = worldwindow.WorldWindow(self, scene)
            self.worldwindow.view.tellarea.connect(self.maparea)
            self.worldwindow.show()
            self.showArea('init')

    def showArea(self, event):
        if self.ignore:
            return
        if self.sender() == self.southSpin or self.sender() == self.eastSpin:
            if self.southSpin.value() > self.northSpin.value():
                y = self.northSpin.value()
                self.northSpin.setValue(self.southSpin.value())
                self.southSpin.setValue(y)
            if self.eastSpin.value() < self.westSpin.value():
                x = self.westSpin.value()
                self.westSpin.setValue(self.eastSpin.value())
                self.eastSpin.setValue(x)
        if self.worldwindow is None:
            return
        approx_area = self.worldwindow.view.drawRect([QtCore.QPointF(self.westSpin.value(), self.northSpin.value()),
                                   QtCore.QPointF(self.eastSpin.value(), self.southSpin.value())])
        self.approx_area.setText(approx_area)
        if event != 'init':
            self.worldwindow.view.statusmsg.emit(approx_area)

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    if len(sys.argv) > 1 and sys.argv[1][-4:] != '.ini':
        if not len(sys.argv) >= 6 or sys.argv[1] == '?':
            raise SystemExit('Usage: north_lat west_lon south_lat east_lon output_file zoom=zoom ' +
                             'width=width height=height url=map_url')
        upper_lat = float(sys.argv[1])
        upper_lon = float(sys.argv[2])
        lower_lat = float(sys.argv[3])
        lower_lon = float(sys.argv[4])
        output = sys.argv[5]
        zoom = 6
        width = None
        height = None
        url = None
        debug = None
        if len(sys.argv) > 5:
            try:
                zoom = int(sys.argv[6])
            except:
                pass
            for arg in sys.argv[5 :]:
                if arg[:5] == 'zoom=':
                    try:
                        zoom = int(arg[5:])
                    except:
                        pass
                elif arg[:6] == 'width=':
                    try:
                        width = int(arg[6:])
                    except:
                        pass
                elif arg[:7] == 'height=':
                    try:
                        height = int(arg[7:])
                    except:
                        pass
                elif arg[:4] == 'url=':
                    try:
                        url = arg[4:]
                    except:
                        pass
                elif arg[:6] == 'debug=':
                    try:
                        if arg[6:7].lower() in ['y', 't', '1']:
                            debug = True
                    except:
                        pass

        retrieveMap(upper_lat, upper_lon, lower_lat, lower_lon, output, zoom=zoom, width=width, height=height, url=url, debug=debug)
    else:
        ex = getMap()
        app.exec_()
        app.deleteLater()
        sys.exit()
