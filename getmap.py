#!/usr/bin/python
#
#  Copyright (C) 2015-2016 Sustainable Energy Now Inc., Angus King
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

import httplib
import math
import os
import sys
import tempfile
from PyQt4 import QtGui, QtCore
import displayobject
from credits import fileVersion

scale = {0: '1:500 million', 1: '1:250 million', 2: '1:150 million', 3: '1:70 million',
         4: '1:35 million', 5: '1:15 million', 6: '1:10 million', 7: '1:4 million',
         8: '1:2 million', 9: '1:1 million', 10: '1:500,000', 11: '1:250,000',
         12: '1:150,000', 13: '1:70,000', 14: '1:35,000', 15: '1:15,000', 16: '1:8,000',
         17: '1:4,000', 18: '1:2,000', 19: '1:1,000'}


class GetMany(QtGui.QDialog):
    def __init__(self, parent=None):
        super(GetMany, self).__init__(parent)
        layout = QtGui.QVBoxLayout(self)
        layout.addWidget(QtGui.QLabel('Enter Coordinates'))
        self.text = QtGui.QPlainTextEdit()
        self.text.setPlainText('Enter list of coordinates separated by spaces or commas. west lat.,' \
                               + ' north lon., east lat., south lon. ...')
        layout.addWidget(self.text)
         # OK and Cancel buttons
        buttons = QtGui.QDialogButtonBox(
            QtGui.QDialogButtonBox.Ok | QtGui.QDialogButtonBox.Cancel,
            QtCore.Qt.Horizontal, self)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.setWindowTitle('SIREN getmap (' + fileVersion() + ") - List of Coordinates")

    def list(self):
        coords = str(self.text.toPlainText())
        if coords != '':
            if coords.find(' ') >= 0:
                if coords.find(',') < 0:
                    coords = coords.replace(' ', ',')
                else:
                    coords = coords.replace(' ', '')
            coords = coords.replace('\n', '')
            bits = coords.split(',')
            grids = []
            c = 4
            for i in range(len(bits)):
                if bits[i].lstrip('-').replace('.','',1).isdigit():
                    if c >= 3:
                        grids.append([])
                        c = -1
                    c += 1
                    try:
                        grids[-1].append(float(bits[i]))
                    except:
                        grids[-1].append(0.)
            return grids
        return None

     # static method to create the dialog and return
    @staticmethod
    def getList(parent=None):
        dialog = GetMany(parent)
        result = dialog.exec_()
        return (dialog.list())


class retrieveMap():

    def numTiles(self, z):
         return(float(pow(2, z)))

    def mercatorToLat(self, mercatorY):
        return(math.degrees(math.atan(math.sinh(mercatorY))))

    def latEdges(self, y, z):
        n = self.numTiles(z)
        unit = 1 / n
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
        n = self.numTiles(z)
        unit = 360 / n
        lon1 = -180 + x * unit
        lon2 = lon1 + unit
        return(lon1, lon2)

    def tileEdges(self, x, y, z):
        lat1, lat2 = self.latEdges(y, z)
        lon1, lon2 = self.lonEdges(x, z)
        return((lat2, lon1, lat1, lon2))   # S,W,N,E

    def deg2num(self, lat_deg, lon_deg, zoom):
        lat_rad = math.radians(lat_deg)
        n = 2.0 ** zoom
        xtile = int((lon_deg + 180.0) / 360.0 * n)
        try:
            ytile = int((1.0 - math.log(math.tan(lat_rad) + (1 / math.cos(lat_rad))) / math.pi) / 2.0 * n)
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
        if len(sys.argv) > 1:
            print url + url_tail
        conn = httplib.HTTPConnection(url)
        if len(sys.argv) > 1:
            print 'Retrieving ' + url_tail
      #   conn.addheaders = [('User-agent', 'Mozilla/5.0')]
        conn.request('GET', url_tail)
        response = conn.getresponse()
        if response.status == 200 and response.reason == 'OK':
            if len(sys.argv) > 1:
                print url_tail + ' retrieved'
            f = open(self.tmp_location + file_name, 'wb')
            f.write(response.read())
            f.close()
        else:
            if len(sys.argv) > 1:
                print url_tail + ' failed'
                print str(response.status) + ' ' + response.reason
        conn.close()
        return file_name

    def __init__(self, upper_lat, upper_lon, lower_lat, lower_lon, output, zoom=None, url=None, width=None, height=None):
        self.log = ''
        self.properties = ''
        if width != None and height != None: # Mapquest map
            url = 'www.mapquestapi.com'
            url_tail = '/staticmap/v4/getmap?type=sat&margin=0&' + \
                       'bestfit=%s,%s,%s,%s&size=%s,%s&imagetype=%s' % (upper_lat, upper_lon, lower_lat, lower_lon,
                       width, height, output[output.rfind('.') + 1:])
            url_key = '&key=yWspjYHSK6FHtNLzZVcqP3WBxSWSwEo8'
            if len(sys.argv) > 1:
                print url + url_tail
            conn = httplib.HTTPConnection(url)
            if len(sys.argv) > 1:
                print 'Requesting ' + url_tail
            conn.request('GET', url_tail + url_key)
            response = conn.getresponse()
            if response.status == 200 and response.reason == 'OK':
                if len(sys.argv) > 1:
                    print url_tail + ' retrieved'
                f = open(output, 'wb')
                f.write(response.read())
                f.close()
                self.log += '\nSaving map to ' + output
                self.properties = 'map_choice=?'
                self.properties += '\nmap=%s' % (output)
                self.properties += '\nupper_left=%1.3f, %1.3f' % (upper_lat, upper_lon)
                self.properties += '\nlower_right=%1.3f, %1.3f' % (lower_lat, lower_lon)
            else:
                if len(sys.argv) > 1:
                    print url_tail + ' failed'
                    print str(response.status) + ' ' + response.reason
            conn.close()
            return
        if url is None:
            self.url = 'http://[abc].tile.openstreetmap.org/zoom/x/y.png'
        else:
            self.url = url
        top_left = self.deg2num(upper_lat, upper_lon, zoom)
        bottom_right = self.deg2num(lower_lat, lower_lon, zoom)
        height = (bottom_right[1] - top_left[1] + 1) * 256
        width = (bottom_right[0] - top_left[0] + 1) * 256
        st, wt, nt, et = self.tileEdges(top_left[0], top_left[1], zoom)
        sb, wb, nb, eb = self.tileEdges(bottom_right[0], bottom_right[1], zoom)
        if len(sys.argv) > 1:
            print '(124)', '%d: %d,%d --> %1.3f :: %1.3f, %1.3f :: %1.3f' % (zoom, top_left[0], top_left[1], st, nt, wt, et)
            print '(125)', '%d: %d,%d --> %1.3f :: %1.3f, %1.3f :: %1.3f' % (zoom, bottom_right[0], bottom_right[1], sb, nb, wb, eb)
        w = bottom_right[0] - top_left[0] + 1
        h = bottom_right[1] - top_left[1] + 1
        if len(sys.argv) > 1:
            print w, h, '=', w * h, 'tiles.', w * 256, 'x', h * 256, 'pixels (approx.', w * 256 * h * 256, 'uncompressed bytes)'
            print 'map_choice=%s' % (zoom)
            print 'map%s=%s' % (zoom, output)
            print 'upper_left%d=%1.3f, %1.3f' % (zoom, nt, wt)
            print 'lower_right%d=%1.3f, %1.3f' % (zoom, sb, eb)
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
        if len(sys.argv) > 1:
            print 'Saving map to ' + fname
        else:
            self.log += '\nSaving map to ' + fname
        for x in range(top_left[0], bottom_right[0] + 1):
            for y in range(top_left[1], bottom_right[1] + 1):
                foo = QtGui.QImage(self.tmp_location + self.writetile(x, y, zoom))
                painter.drawImage(QtCore.QPoint(256 * (x - top_left[0]), 256 * (y - top_left[1])), foo)
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


class ClickableQLabel(QtGui.QLabel):
    def __init(self, parent):
        QLabel.__init__(self, parent)

    def mousePressEvent(self, event):
        self.emit(QtCore.SIGNAL('clicked()'))


class getMap(QtGui.QWidget):

    def deg2num(self, lat_deg, lon_deg):
        if lat_deg > 85:
            ld = 85.
        elif lat_deg < -85:
            ld = -85.
        else:
            ld = lat_deg
        lat_rad = math.radians(ld)
        n = self.world_width
        xtile = (lon_deg + 180.0) / 360.0 * n
        try:
            ytile = (1.0 - math.log(math.tan(lat_rad) + (1 / math.cos(lat_rad))) / math.pi) / 2.0 * n
        except:
            ytile = 0
        return (xtile, ytile)

    def __init__(self, help='help.html'):
        super(getMap, self).__init__()
        self.help = help
        self.initUI()

    def initUI(self):
        north = QtGui.QLabel('North')
        south = QtGui.QLabel('South')
        east = QtGui.QLabel('East')
        west = QtGui.QLabel('West')
        self.northSpin = QtGui.QDoubleSpinBox()
        self.northSpin.setDecimals(3)
        self.northSpin.setSingleStep(5)
        self.northSpin.setRange(-85, 85)
        self.northSpin.setValue(-30.)
        self.southSpin = QtGui.QDoubleSpinBox()
        self.southSpin.setDecimals(3)
        self.southSpin.setSingleStep(5)
        self.southSpin.setRange(-85, 85)
        self.southSpin.setValue(-35.)
        self.westSpin = QtGui.QDoubleSpinBox()
        self.westSpin.setDecimals(3)
        self.westSpin.setSingleStep(5)
        self.westSpin.setRange(-180, 180)
        self.westSpin.setValue(115)
        self.eastSpin = QtGui.QDoubleSpinBox()
        self.eastSpin.setDecimals(3)
        self.eastSpin.setSingleStep(5)
        self.eastSpin.setRange(-180, 180)
        self.eastSpin.setValue(120)
        zoom = QtGui.QLabel('Map Scale (Zoom):')
        self.zoomSpin = QtGui.QSpinBox()
        self.zoomSpin.setValue(6)
        self.zoomSpin.setRange(0, 11)
        self.zoomSpin.valueChanged[str].connect(self.zoomChanged)
        self.zoomScale = QtGui.QLabel('(' + scale[6] + ')')
        self.grid = QtGui.QGridLayout()
        self.grid.addWidget(QtGui.QLabel('Area of Interest:'), 0, 0)
        self.grid.addWidget(north, 0, 1)
        self.world = False
        if os.path.exists('world1.jpg'):
            self.world = 'world1.jpg'
            a_nl = ''
        elif os.path.exists('world.jpg'):
            self.world = 'world.jpg'
            a_nl = '\n'
        if self.world:
            self.do_world = True
            self.wmap = QtGui.QLabel()
            world = QtGui.QImage(self.world)
            self.world_width = world.width()
            self.world_height = world.height()
            self.width_ratio = world.width() / 360.
            self.height_ratio = world.height() / 170.1022
            painter = QtGui.QPainter()
            painter.begin(world)
            painter.setPen(QtGui.QPen(QtGui.QBrush(QtCore.Qt.red), 0))
            painter.drawRect(1., 2., 50., 50.)
            painter.end()
            self.wmap.setPixmap(QtGui.QPixmap.fromImage(world))
            self.grid.addWidget(self.wmap, 1, 2)
        else:
            self.do_world = False
            self.world_width = 100
            self.world_height = 0
        self.grid.setAlignment(north, QtCore.Qt.AlignRight)
        self.grid.addWidget(self.northSpin, 0, 2)
        self.grid.addWidget(west, 1, 0)
        self.grid.setAlignment(west, QtCore.Qt.AlignRight)
        self.grid.addWidget(self.westSpin, 1, 1)
        self.grid.addWidget(east, 1, 3)
        self.grid.setAlignment(east, QtCore.Qt.AlignRight)
        self.grid.addWidget(self.eastSpin, 1, 4)
        self.grid.addWidget(south, 2, 1)
        self.grid.setAlignment(south, QtCore.Qt.AlignRight)
        self.grid.addWidget(self.southSpin, 2, 2)
        self.grid.addWidget(zoom, 3, 0)
        self.grid.addWidget(self.zoomSpin, 3, 1)
        self.grid.addWidget(self.zoomScale, 3, 2)
        self.grid.addWidget(QtGui.QLabel('URL template:'), 4, 0)
        self.urltemplate = QtGui.QLineEdit()
        self.urltemplate.setText('http://[abc].tile.openstreetmap.org/zoom/x/y.png')
        self.grid.addWidget(self.urltemplate, 4, 1, 1, 5)
        self.grid.addWidget(QtGui.QLabel('Image Width:'), 5, 0)
        self.widthSpin = QtGui.QSpinBox()
        self.widthSpin.setSingleStep(50)
        self.widthSpin.setRange(50, 3840)
        self.widthSpin.setValue(400)
        self.grid.addWidget(self.widthSpin, 5, 1)
        hlabel = QtGui.QLabel('Image Height:   ')
        hlabel.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
        self.grid.addWidget(hlabel, 5, 2)
        self.heightSpin = QtGui.QSpinBox()
        self.heightSpin.setSingleStep(50)
        self.heightSpin.setRange(50, 3840)
        self.heightSpin.setValue(200)
        self.grid.addWidget(self.heightSpin, 5, 3)
        self.grid.addWidget(QtGui.QLabel('Image File name:'), 6, 0)
        cur_dir = os.getcwd()
        self.filename = ClickableQLabel()
        self.filename.setText(cur_dir + '/untitled.png')
        self.filename.setFrameStyle(6)
        self.connect(self.filename, QtCore.SIGNAL('clicked()'), self.fileChanged)
        self.grid.addWidget(self.filename, 6, 1, 1, 5)
        self.grid.addWidget(QtGui.QLabel('Properties:'), 7, 0)
        self.properties = QtGui.QPlainTextEdit()
        self.properties.setMaximumHeight(north.sizeHint().height() * 4.5)
        self.properties.setReadOnly(True)
        self.grid.addWidget(self.properties, 7, 1, 3, 5)
        self.log = QtGui.QLabel()
        self.grid.addWidget(self.log, 10, 1, 3, 5)
        quit = QtGui.QPushButton('Quit', self)
        self.grid.addWidget(quit, 14, 0)
        quit.clicked.connect(self.quitClicked)
        QtGui.QShortcut(QtGui.QKeySequence('q'), self, self.quitClicked)
        query = QtGui.QPushButton('Query Map', self)
        wdth = query.fontMetrics().boundingRect(query.text()).width() + 9
        self.grid.addWidget(query, 14, 1)
        query.clicked.connect(self.queryClicked)
        make = QtGui.QPushButton('Get Map', self)
        make.setMaximumWidth(wdth)
        self.grid.addWidget(make, 14, 2)
        make.clicked.connect(self.makeClicked)
        mapquest = QtGui.QPushButton('MapQuest', self)
        mapquest.setMaximumWidth(wdth)
        self.grid.addWidget(mapquest, 14, 3)
        mapquest.clicked.connect(self.mapquestClicked)
        help = QtGui.QPushButton('Help', self)
        help.setMaximumWidth(wdth)
        self.grid.addWidget(help, 14, 4)
        help.clicked.connect(self.helpClicked)
        QtGui.QShortcut(QtGui.QKeySequence('F1'), self, self.helpClicked)
        note = QtCore.QString('Map data ' + unichr(169) + ' OpenStreetMap contributors CC-BY-SA ' +
               a_nl + '(http://www.openstreetmap.org/copyright)')
        self.grid.addWidget(QtGui.QLabel(note), 15, 0, 1, 3)
        many = QtGui.QPushButton('Many', self)
        many.setMaximumWidth(wdth)
        self.grid.addWidget(many, 15, 4)
        many.clicked.connect(self.manyClicked)
        saveView = QtGui.QPushButton('Save View', self)
        saveView.setMaximumWidth(wdth)
        self.grid.addWidget(saveView, 15, 3)
        saveView.clicked.connect(self.saveViewClicked)
        frame = QtGui.QFrame()
        frame.setLayout(self.grid)
        self.scroll = QtGui.QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(frame)
        self.layout = QtGui.QVBoxLayout(self)
        self.layout.addWidget(self.scroll)
        self.resize(self.width() + int(self.world_width * .7), self.height() + int(self.world_height * .7))
        self.setWindowTitle('SIREN getmap (' + fileVersion() + ") - Make Map from OSM or MapQuest")
        self.center()
        self.show()

    def center(self):
        frameGm = self.frameGeometry()
        screen = QtGui.QApplication.desktop().screenNumber(QtGui.QApplication.desktop().cursor().pos())
        centerPoint = QtGui.QApplication.desktop().availableGeometry(screen).center()
        frameGm.moveCenter(centerPoint)
        self.move(frameGm.topLeft())

    def zoomChanged(self, val):
        self.zoomScale.setText('(' + scale[int(val)] + ')')
        self.zoomScale.adjustSize()

    def fileChanged(self):
        self.filename.setText(QtGui.QFileDialog.getSaveFileName(self, 'Save Image File',
                              self.filename.text(),
                              'Images (*.jpeg *.jpg *.png)'))
        if self.filename.text() != '':
            i = str(self.filename.text()).rfind('.')
            if i < 0:
                self.filename.setText(self.filename.text() + '.png')

    def helpClicked(self):
        dialog = displayobject.AnObject(QtGui.QDialog(), self.help,
                 title='Help for SIREN getmap (' + fileVersion() + ')', section='map')
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
               '?', self.zoomSpin.value(), url=str(self.urltemplate.text()))
        self.properties.setPlainText(mapp.getProperties())
        self.log.setText(mapp.getLog())
        if self.do_world:
            coords = mapp.getCoords()
            x, y = self.deg2num(coords[0], coords[1])
            x2, y2 = self.deg2num(coords[2], coords[3])
            h = y2 - y
            w = x2 - x
            world = QtGui.QImage(self.world)
            painter = QtGui.QPainter()
            painter.begin(world)
            painter.setPen(QtGui.QPen(QtGui.QBrush(QtCore.Qt.red), 0))
            painter.drawRect(x, y, w, h)
            painter.end()
            self.wmap.setPixmap(QtGui.QPixmap.fromImage(world))

    def manyClicked(self):
        if self.do_world:
            grids = GetMany.getList()
            if grids is None:
                return
            world = QtGui.QImage(self.world)
            painter = QtGui.QPainter()
            painter.begin(world)
            for g in range(len(grids)):
                mapp = retrieveMap(grids[g][0], grids[g][1], grids[g][2], grids[g][3], '?',
                       self.zoomSpin.value(), url=str(self.urltemplate.text()))
                coords = mapp.getCoords()
                x, y = self.deg2num(coords[0], coords[1])
                x2, y2 = self.deg2num(coords[2], coords[3])
                h = y2 - y
                w = x2 - x
                painter.setPen(QtGui.QPen(QtGui.QBrush(QtCore.Qt.red), 0))
                painter.drawRect(x, y, w, h)
            painter.end()
            self.wmap.setPixmap(QtGui.QPixmap.fromImage(world))

    def saveViewClicked(self):
        outputimg = self.wmap.pixmap()
        fname = 'getmap_view.png'
        fname = QtGui.QFileDialog.getSaveFileName(self, 'Save image file',
                fname, 'Image Files (*.png *.jpg *.bmp)')
        if fname != '':
            fname = str(fname)
            i = fname.rfind('.')
            if i < 0:
                fname = fname + '.png'
                i = fname.rfind('.')
            outputimg.save(fname, fname[i + 1:])

    def makeClicked(self):
        mapp = retrieveMap(self.northSpin.value(), self.westSpin.value(), self.southSpin.value(), self.eastSpin.value(),
               str(self.filename.text()), zoom=self.zoomSpin.value(), url=str(self.urltemplate.text()))
        self.properties.setPlainText(mapp.getProperties())
        self.log.setText(mapp.getLog())

    def mapquestClicked(self):
        mapp = retrieveMap(self.northSpin.value(), self.westSpin.value(), self.southSpin.value(), self.eastSpin.value(),
               str(self.filename.text()), width=self.widthSpin.value(), height=self.heightSpin.value())
        self.properties.setPlainText(mapp.getProperties())
        self.log.setText(mapp.getLog())


if __name__ == '__main__':
    app = QtGui.QApplication(sys.argv)
    if len(sys.argv) > 1:
        if not len(sys.argv) > 6:
            raise SystemExit('Usage: north_lat west_lon south_lat east_lon output_file zoom=zoom' +
                             'width=width height=height url=map_url')
        upper_lat = float(sys.argv[1])
        upper_lon = float(sys.argv[2])
        lower_lat = float(sys.argv[3])
        lower_lon = float(sys.argv[4])
        zoom = int(sys.argv[5])
        output = sys.argv[6]
        if len(sys.argv) > 7:
            url = sys.argv[7]
            retrieveMap(upper_lat, upper_lon, lower_lat, lower_lon, output, zoom=zoom, url=url)
        else:
            retrieveMap(upper_lat, upper_lon, lower_lat, lower_lon, output, zoom=zoom)
    else:
        ex = getMap()
        app.exec_()
        app.deleteLater()
        sys.exit()
