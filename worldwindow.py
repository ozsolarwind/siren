#!/usr/bin/python3
#
#  Copyright (C) 2017-2019 Sustainable Energy Now Inc., Angus King
#
#  worldwindow.py - This file is part of SIREN.
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

from math import asin, atan2, cos, degrees, pi, pow, radians, sin, sqrt
import os
import sys

import configparser   # decode getfiles.ini file
from PyQt5 import QtCore, QtGui, QtWidgets

#try:
#    from mpl_toolkits.basemap.pyproj import Proj  # Import the pyproj.Proj module
#except:
from mpl_toolkits.basemap import pyproj as pyproj
from pyproj import Proj as Proj

from colours import Colours
from credits import fileVersion
import displayobject
from editini import EdtDialog, EditSect, SaveIni
from parents import getParents
from senuser import getUser

RADIUS = 6367.

scale = {0: '1:500 million', 1: '1:250 million', 2: '1:150 million', 3: '1:70 million',
         4: '1:35 million', 5: '1:15 million', 6: '1:10 million', 7: '1:4 million',
         8: '1:2 million', 9: '1:1 million', 10: '1:500,000', 11: '1:250,000',
         12: '1:150,000', 13: '1:70,000', 14: '1:35,000', 15: '1:15,000', 16: '1:8,000',
         17: '1:4,000', 18: '1:2,000', 19: '1:1,000'}

def p2str(p):
    return '(%.4f,%.4f)' % (p.y(), p.x())

def reproject(latitude, longitude):
    """Returns the x & y coordinates in meters using a sinusoidal projection"""
    earth_radius = RADIUS # in meters
    lat_dist = pi * earth_radius / 180.0

    y = [lat * lat_dist for lat in latitude]
    x = [int * lat_dist * cos(radians(lat))
                for lat, int in zip(latitude, longitude)]
    return x, y

def area_of_polygon(x, y):
    """Calculates the area of an arbitrary polygon given its verticies"""
    area = 0.0
    for i in range(-1, len(x)-1):
        area += x[i] * (y[i+1] - y[i-1])
    return abs(area) / 2.0

def merra_cells(top, lft, bot, rht):
    top = round((top + 0.25) * 2) / 2
    bot = round((bot - 0.25) * 2) / 2
    lat = (top - bot) * 2
    lft = round((lft - 0.3125) / .625) * .625
    rht = round((rht + 0.3125) / .625) * .625
    lon = (rht - lft) / 0.625
    return 'Lat: %s x Lon: %s' % (str(int((top - bot) * 2)), str(int((rht - lft) / 0.625)))


class GetMany(QtWidgets.QDialog):

    def __init__(self, parent=None):
        super(GetMany, self).__init__(parent)
        layout = QtWidgets.QVBoxLayout(self)
        layout.addWidget(QtWidgets.QLabel('Enter Coordinates'))
        self.text = QtWidgets.QPlainTextEdit()
        self.text.setPlainText('Enter list of coordinates separated by spaces or commas. west lat.,' \
                               + ' north lon., east lat., south lon. ...')
        layout.addWidget(self.text)
         # OK and Cancel buttons
        buttons = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel,
            QtCore.Qt.Horizontal, self)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.setWindowTitle('SIREN - worldwindow (' + fileVersion() + ") - List of Coordinates")
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))

    def list(self):
        coords = self.text.toPlainText()
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


class WorldScene(QtWidgets.QGraphicsScene):

    def get_config(self):
        config = configparser.RawConfigParser()
        config_file = 'getfiles.ini'
        config.read(config_file)
        parents = []
        try:
            parents = getParents(config.items('Parents'))
        except:
            pass
        try:
            self.map = config.get('Map', 'map_choice')
        except:
            self.map = ''
        try:
            self.map_file = config.get('Map', 'map' + self.map)
            for key, value in parents:
                self.map_file = self.map_file.replace(key, value)
            self.map_file = self.map_file.replace('$USER$', getUser())
        except:
            self.map_file = ''
        try:
            self.projection = config.get('Map', 'projection' + self.map)
        except:
            try:
                self.projection = config.get('Map', 'projection')
            except:
                self.projection = 'EPSG:3857'
        self.map_upper_left = [85.06, -180.]
        self.map_lower_right = [-85.06, 180.]
        try:
             upper_left = config.get('Map', 'upper_left' + self.map).split(',')
             self.map_upper_left[0] = float(upper_left[0].strip())
             self.map_upper_left[1] = float(upper_left[1].strip())
             lower_right = config.get('Map', 'lower_right' + self.map).split(',')
             self.map_lower_right[0] = float(lower_right[0].strip())
             self.map_lower_right[1] = float(lower_right[1].strip())
        except:
             try:
                 lower_left = config.get('Map', 'lower_left' + self.map).split(',')
                 upper_right = config.get('Map', 'upper_right' + self.map).split(',')
                 self.map_upper_left[0] = float(upper_right[0].strip())
                 self.map_upper_left[1] = float(lower_left[1].strip())
                 self.map_lower_right[0] = float(lower_left[0].strip())
                 self.map_lower_right[1] = float(upper_right[1].strip())
             except:
                 pass
        self.map_polygon = [[self.map_upper_left[0], self.map_upper_left[1]],
                            [self.map_upper_left[0], self.map_lower_right[1]],
                            [self.map_lower_right[0], self.map_lower_right[1]],
                            [self.map_lower_right[0], self.map_upper_left[1]],
                            [self.map_upper_left[0], self.map_upper_left[1]]]
        self.colors = {}
        self.colors['background'] = 'darkBlue'
        self.colors['border'] = 'red'
        self.colors['grid'] = 'gray'
        try:
            colours = config.items('Colors')
            for item, colour in colours:
                self.colors[item] = colour
        except:
            pass
        if self.map != '':
            try:
                colours = config.items('Colors' + self.map)
                for item, colour in colours:
                    self.colors[item] = colour
            except:
                pass
        self.show_ruler = False
        try:
            show_ruler = config.get('View', 'show_ruler')
            if show_ruler.lower() in ['true', 'yes', 'on']:
                self.show_ruler = True
        except:
            pass
        self.ruler = 1000.
        self.ruler_ticks = 100.
        try:
            ruler = config.get('View', 'ruler')
            rule = ruler.split(',')
            self.ruler = float(rule[0])
            if len(rule) > 1:
                self.ruler_ticks = float(rule[1].strip())
        except:
            pass
        self.center_on_grid = False
        try:
            center_on_grid = config.get('View', 'center_on_grid')
            if center_on_grid.lower() in ['true', 'yes', 'on']:
                self.center_on_grid = True
        except:
            pass
        self.line_width = 0
        try:
            line_width = config.get('View', 'line_width')
            try:
                self.line_width = int(line_width)
            except:
                self.line_width = float(line_width)
        except:
            pass
        self.show_grid = False
        try:
            show_grid = config.get('View', 'show_grid')
            if show_grid.lower() in ['true', 'yes', 'on']:
                self.show_grid = True
        except:
            pass
        self.show_mgrid = False
        try:
            show_mgrid = config.get('View', 'show_merra_grid')
            if show_mgrid.lower() in ['true', 'yes', 'on']:
                self.show_mgrid = True
        except:
            pass

    def destinationxy(self, lon1, lat1, bearing, distance):
        """
        Given a start point, initial bearing, and distance, calculate
        the destination point and final bearing travelling along a
        (shortest distance) great circle arc
        """
        radius = RADIUS   # km is the radius of the Earth
     # convert decimal degrees to radians
        ln1, lt1, baring = list(map(radians, [lon1, lat1, bearing]))
     # "reverse" haversine formula
        lat2 = asin(sin(lt1) * cos(distance / radius) +
                                cos(lt1) * sin(distance / radius) * cos(baring))
        lon2 = ln1 + atan2(sin(baring) * sin(distance / radius) * cos(lt1),
                                            cos(distance / radius) - sin(lt1) * sin(lat2))
        return degrees(lon2), degrees(lat2)

    def __init__(self):
        QtWidgets.QGraphicsScene.__init__(self)
        self.exitLoop = False
        self.loopMax = 0
        self.get_config()
        self._gridGroup = None
        self._mgridGroup = None
        self.setBackgroundBrush(QtGui.QColor(self.colors['background']))
        if os.path.exists(self.map_file):
            war = QtGui.QImageReader(self.map_file)
            wa = war.read()
            pixMap = QtGui.QPixmap.fromImage(wa)
        else:
            ratio = abs(self.map_upper_left[1] - self.map_lower_right[1]) / \
                     abs(self.map_upper_left[0] - self.map_lower_right[0])
            pixMap = QtGui.QPixmap(200 * ratio, 200)
            if self.map != '':
                painter = QtGui.QPainter(pixMap)
                brush = QtGui.QBrush(QtGui.QColor('white'))
                painter.fillRect(QtCore.QRectF(0, 0, 200 * ratio, 200), brush)
                painter.setPen(QtGui.QColor('lightgray'))
                painter.drawText(QtCore.QRectF(0, 0, 200 * ratio, 200), QtCore.Qt.AlignCenter,
                                 'Map not found.')
                painter.end()
        self.addPixmap(pixMap)
        w = self.width()
        h = self.height()
        if isinstance(self.line_width, float):
            if self.line_width < 1:
                self.line_width = w * self.line_width
        self.upper_left = [0, 0, self.map_upper_left[1], self.map_upper_left[0]]
        self.lower_right = [w, h, self.map_lower_right[1], self.map_lower_right[0]]
        self.setSceneRect(-w * 0.05, -h * 0.05, w * 1.1, h * 1.1)
        self._positions = {}
        self._setupCoordTransform()

    def _setupCoordTransform(self):
        self._proj = Proj('+init=' + self.projection)   # LatLon with WGS84 datum used by GPS units and Google Earth
        x1, y1, lon1, lat1 = self.upper_left
        x2, y2, lon2, lat2 = self.lower_right
        ul = self._proj(lon1, lat1)
        lr = self._proj(lon2, lat2)
        self._lat_scale = y2 / (lr[1] - ul[1])
        self._lon_scale = x2 / (lr[0] - ul[0])
        self._orig_lat = ul[1]
        self._orig_lon = ul[0]

    def mapToLonLat(self, p):
        x = p.x() / self._lon_scale + self._orig_lon
        y = p.y() / self._lat_scale + self._orig_lat
        lon, lat = self._proj(x, y, inverse=True)
        return QtCore.QPointF(round(lon, 4), round(lat, 4))

    def mapFromLonLat(self, p):
        lon, lat = p.x(), p.y()
        x, y = self._proj(lon, lat)
        x = (x - self._orig_lon) * self._lon_scale
        y = (y - self._orig_lat) * self._lat_scale
        return QtCore.QPointF(x, y)

    def positions(self):
        try:
            return self._positions
        except:
            return


class WorldView(QtWidgets.QGraphicsView):
    statusmsg = QtCore.pyqtSignal(str)
    tellarea = QtCore.pyqtSignal(list, str)

    def __init__(self, scene, zoom=.8):
        QtWidgets.QGraphicsView.__init__(self, scene)
        self.zoom = zoom
        QtWidgets.QShortcut(QtGui.QKeySequence('pgdown'), self, self.zoomIn)
        QtWidgets.QShortcut(QtGui.QKeySequence('pgup'), self, self.zoomOut)
        QtWidgets.QShortcut(QtGui.QKeySequence('Ctrl++'), self, self.zoomIn)
        QtWidgets.QShortcut(QtGui.QKeySequence('Ctrl+-'), self, self.zoomOut)
        self._rectangle = None
        self._drag_start = None
        self.setMouseTracking(True)
        self.ruler_items = []
        self.legend_items = []
        self.legend_pos = None
        self.rect_items = []

    def clear_Rect(self):
        try:
            for i in range(len(self.rect_items)):
                self.scene().removeItem(self.rect_items[i])
            del self.rect_items
        except:
            pass

    def drawRect(self, coords=None):
        self.clear_Rect()
        self.rect_items = []
        color = QtGui.QColor()
        color.setNamedColor((self.scene().colors['border']))
        pen = QtGui.QPen(color, self.scene().line_width)
        pen.setJoinStyle(QtCore.Qt.RoundJoin)
        pen.setCapStyle(QtCore.Qt.RoundCap)
        fromm = self.mapFromLonLat(coords[0])
        too = self.mapFromLonLat(coords[1])
        self.rect_items.append(QtWidgets.QGraphicsRectItem(fromm.x(), fromm.y(),
            too.x() - fromm.x(), too.y() - fromm.y()))
        self.rect_items[-1].setPen(pen)
        self.rect_items[-1].setZValue(3)
        self.scene().addItem(self.rect_items[-1])
        if self.scene().center_on_grid:
            x = fromm.x() + (too.x() - fromm.x()) / 2
            y = fromm.y() + (too.y() - fromm.y()) / 2
            go_to = QtCore.QPointF(x, y)
            self.centerOn(go_to)
        rect = reproject([coords[0].y(), coords[0].y(), coords[1].y(), coords[1].y()],
                         [coords[0].x(), coords[1].x(), coords[1].x(), coords[0].x()])
        return '{:0,.0f}'.format(area_of_polygon(rect[0], rect[1])) + ' Km^2'

    def drawMany(self):
        grids = GetMany.getList()
        if grids is None:
            return
        self.clear_Rect()
        self.rect_items = []
        color = QtGui.QColor()
        color.setNamedColor((self.scene().colors['border']))
        pen = QtGui.QPen(color, self.scene().line_width)
        pen.setJoinStyle(QtCore.Qt.RoundJoin)
        pen.setCapStyle(QtCore.Qt.RoundCap)
        for g in range(len(grids)):
            fromm = self.mapFromLonLat(QtCore.QPointF(grids[g][1], grids[g][0]))
            too = self.mapFromLonLat(QtCore.QPointF(grids[g][3], grids[g][2]))
            self.rect_items.append(QtWidgets.QGraphicsRectItem(fromm.x(), fromm.y(),
                                   too.x() - fromm.x(), too.y() - fromm.y()))
            self.rect_items[-1].setPen(pen)
            self.rect_items[-1].setZValue(3)
            self.scene().addItem(self.rect_items[-1])
        self.statusmsg.emit(str(len(grids)) + ' Areas drawn')

    def destinationxy(self, lon1, lat1, bearing, distance):
        """
        Given a start point, initial bearing, and distance, calculate
        the destination point and final bearing travelling along a
        (shortest distance) great circle arc
        """
        radius = RADIUS  # km is the radius of the Earth
     # convert decimal degrees to radians
        ln1, lt1, baring = list(map(radians, [lon1, lat1, bearing]))
     # "reverse" haversine formula
        lat2 = asin(sin(lt1) * cos(distance / radius) +
               cos(lt1) * sin(distance / radius) * cos(baring))
        lon2 = ln1 + atan2(sin(baring) * sin(distance / radius) * cos(lt1),
               cos(distance / radius) - sin(lt1) * sin(lat2))
        return QtCore.QPointF(degrees(lon2), degrees(lat2))

    def destinationx(self, lon1, lat1, distance):
        radius = RADIUS  # km is the radius of the Earth
     # convert decimal degrees to radians
        lt1 = radians(lat1)
        circum = radius * 2 * pi * cos(lt1)
        lon2 = lon1 + distance / circum * 360
        return QtCore.QPointF(lon2, lat1)

    def mapToLonLat(self, p):
        p = self.mapToScene(p)
        return self.scene().mapToLonLat(p)

    def mapFromLonLat(self, p):
        p = self.scene().mapFromLonLat(p)
        return p

    def wheelEvent(self, event):
        delta = -event.angleDelta().y() / (100 * (1 + 1 - self.zoom))
        zoom = pow(self.zoom, delta)
        self.scale(zoom, zoom)

    def zoomIn(self):
        self.scale(1 / self.zoom, 1 / self.zoom)

    def zoomOut(self):
        self.scale(self.zoom, self.zoom)

    def mousePressEvent(self, event):
        if QtCore.Qt.LeftButton == event.button():
            where = self.mapToLonLat(event.pos())
            pl = self.mapToLonLat(event.pos())
            self.statusmsg.emit(p2str(pl) + ' ' + p2str(event.pos()))
            if self._rectangle is None:
                self._rectangle = [where]
                hb = self.horizontalScrollBar()
                vb = self.verticalScrollBar()
                self._sb_start = QtCore.QPoint(hb.value(), vb.value())
            else:
                self._rectangle.append(where)
                if self._rectangle[0].x() > self._rectangle[1].x():
                    x = self._rectangle[0].x()
                    self._rectangle[0] = QtCore.QPointF(self._rectangle[1].x(), self._rectangle[0].y())
                    self._rectangle[1] = QtCore.QPointF(x, self._rectangle[1].y())
                if self._rectangle[0].y() < self._rectangle[1].y():
                    y = self._rectangle[0].y()
                    self._rectangle[0] = QtCore.QPointF(self._rectangle[0].x(), self._rectangle[1].y())
                    self._rectangle[1] = QtCore.QPointF(self._rectangle[1].x(), y)
                approx_area = self.drawRect(self._rectangle)
                self.statusmsg.emit(p2str(pl) + ' ' + approx_area)
                self.tellarea.emit(self._rectangle, approx_area)
                self._rectangle = None
            self._drag_start = QtCore.QPoint(event.pos())
            hb = self.horizontalScrollBar()
            vb = self.verticalScrollBar()
            self._sb_start = QtCore.QPoint(hb.value(), vb.value())

    def mouseMoveEvent(self, event):
        if QtCore.Qt.LeftButton == event.buttons() and self._drag_start:
            if self._rectangle:
                self._rectangle = None
            delta = event.pos() - self._drag_start
            val = self._sb_start - delta
            self.horizontalScrollBar().setValue(val.x())
            self.verticalScrollBar().setValue(val.y())
        return
        if QtCore.Qt.LeftButton == event.buttons() and self._rectangle:
            delta = event.pos() - self._rectangle
            val = self._sb_start - delta
            self.horizontalScrollBar().setValue(val.x())
            self.verticalScrollBar().setValue(val.y())
        else:
            p = self.mapToScene(event.pos())
            x, y = list(map(int, [p.x(), p.y()]))
        pl = self.mapToLonLat(event.pos())
      #  self.emit(QtCore.SIGNAL('statusmsg'), p2str(pl) + ' ' + p2str(event.pos()))

    def mouseReleaseEvent(self, event):
        if QtCore.Qt.LeftButton == event.button():
            self._drag_start = None
            self._sb_start = None

    def resizeEvent(self, event):
        self.update()
        self.zoomIn()
        self.zoomOut()

    def show_Ruler(self, ruler, ruler_ticks=None, pos=None):
        def do_ruler():
            if ruler_ticks is None or ruler_ticks <= 0:
                ticks = int(ruler / 5.)
            else:
                ticks = int(ruler / ruler_ticks)
            frll = self.mapToLonLat(p)
            start = self.mapFromLonLat(frll)
            toll = self.destinationxy(frll.x(), frll.y(), 0, ruler)
            end = self.mapFromLonLat(toll)
            self.ruler_items.append(QtWidgets.QGraphicsLineItem(QtCore.QLineF(start, end)))
            self.ruler_items[-1].setPen(pen)
            self.ruler_items[-1].setZValue(0)
            self.scene().addItem(self.ruler_items[-1])
            toll = self.destinationx(frll.x(), frll.y(), ruler)
            end = self.mapFromLonLat(toll)
            self.ruler_items.append(QtWidgets.QGraphicsLineItem(QtCore.QLineF(start, end)))
            self.ruler_items[-1].setPen(pen)
            self.ruler_items[-1].setZValue(0)
            self.scene().addItem(self.ruler_items[-1])
            for i in range(ticks + 1):
                strt = self.destinationx(frll.x(), frll.y(), ruler * i / ticks)
                start = self.mapFromLonLat(strt)
                toll = self.destinationxy(strt.x(), strt.y(), 0, ruler / 50)
                end = self.mapFromLonLat(toll)
                self.ruler_items.append(QtWidgets.QGraphicsLineItem(QtCore.QLineF(start, end)))
                self.ruler_items[-1].setPen(pen)
                self.ruler_items[-1].setZValue(0)
                self.scene().addItem(self.ruler_items[-1])
            self.ruler_items.append(QtWidgets.QGraphicsSimpleTextItem(str(int(ruler)) + ' Km'))
            new_font = self.ruler_items[-1].font()
            new_font.setPointSizeF(self.scene().width() / 90.)
            up = float(QtGui.QFontMetrics(new_font).height())
            self.ruler_items[-1].setFont(new_font)
            self.ruler_items[-1].setPos(end + QtCore.QPointF(5., -up))
            self.ruler_items[-1].setBrush(color)
            self.ruler_items[-1].setZValue(0)
            self.scene().addItem(self.ruler_items[-1])
            for i in range(ticks + 1):
                strt = self.destinationxy(frll.x(), frll.y(), 0, ruler * i / ticks)
                start = self.mapFromLonLat(strt)
                toll = self.destinationx(strt.x(), strt.y(), ruler / 50)
                end = self.mapFromLonLat(toll)
                self.ruler_items.append(QtWidgets.QGraphicsLineItem(QtCore.QLineF(start, end)))
                self.ruler_items[-1].setPen(pen)
                self.ruler_items[-1].setZValue(0)
                self.scene().addItem(self.ruler_items[-1])

        self.ruler_items = []
        if pos is None:
            p = QtCore.QPoint(50, self.height() - 50)
        else:
            p = self.mapFromParent(pos)
        color = QtGui.QColor()
        if self.scene().colors['ruler'].lower() == 'guess':
            sp = self.mapToGlobal(p)
            img = QtGui.QPixmap(5, 5)
            painter = QtGui.QPainter(img)
            targetrect = QtCore.QRectF(0, 0, 5, 5)
            sourcerect = QtCore.QRect(sp.x(), sp.y(), 5, 5)
            self.render(painter, targetrect, sourcerect)
            painter.end()
            ig = img.toImage()
            colorsum = 0
            for x in range(5):
                for y in range(5):
                    c = ig.pixel(x, y)
                    colors = QtGui.QColor(c).getRgbF()
                    colorsum += (colors[0] + colors[1] + colors[2])
            if colorsum <= 50.:   # more black than white
                color.setNamedColor('white')
            else:
                color.setNamedColor('black')
        else:
            color.setNamedColor(self.scene().colors['ruler'])
        pen = QtGui.QPen(color, 0)
        pen.setJoinStyle(QtCore.Qt.RoundJoin)
        pen.setCapStyle(QtCore.Qt.RoundCap)
        do_ruler()

    def hide_Ruler(self):
        for i in range(len(self.ruler_items)):
            self.scene().removeItem(self.ruler_items[i])
        del self.ruler_items


class WorldWindow(QtWidgets.QMainWindow):
    log = QtCore.pyqtSignal()
    statusmsg = QtCore.pyqtSignal(str)
    tellarea = QtCore.pyqtSignal()

    def get_config(self):
        self.config = configparser.RawConfigParser()
        self.config_file = 'getfiles.ini'
        self.config.read(self.config_file)
        try:
            self.help = self.config.get('Files', 'help')
        except:
            self.help = 'help.html'
        self.restorewindows = False
        try:
            rw = self.config.get('Windows', 'restorewindows')
            if rw.lower() in ['true', 'yes', 'on']:
                self.restorewindows = True
        except:
            pass
        self.zoom = 0.8
        try:
            self.zoom = float(self.config.get('View', 'zoom_rate'))
            if self.zoom > 1:
                self.zoom = 1 / self.zoom
            if self.zoom > 0.95:
                self.zoom = 0.95
            elif self.zoom < 0.75:
                self.zoom = 0.75
        except:
            pass

    def __init__(self, parent, scene):
        super(WorldWindow, self).__init__(parent)
        self.get_config()
       #  self.setWindowFlags(QtCore.Qt.FramelessWindowHint)
        self.view = WorldView(scene, self.zoom)
        self.view.scale(1., 1.)
        self._mv = self.view
        self.grid_items = None
        self.setStatusBar(QtWidgets.QStatusBar())
        self.view.statusmsg.connect(self.setStatusText)
        w = QtWidgets.QWidget()
        lay = QtWidgets.QVBoxLayout(w)
        lay.addWidget(self.view)
        self.setCentralWidget(w)
        self.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self.popup)
        menubar = self.menuBar()
        self.showRuler = QtWidgets.QAction(QtGui.QIcon('blank.png'), 'Scale Ruler', self)
        self.showRuler.setShortcut('Ctrl+R')
        self.showRuler.setStatusTip('Show Scale Ruler')
        self.showRuler.triggered.connect(self.show_Ruler)
        self.showGrid = QtWidgets.QAction(QtGui.QIcon('blank.png'), 'Coordinates Grid', self)
        self.showGrid.setShortcut('Ctrl+G')
        self.showGrid.setStatusTip('Show Coordinates Grid')
        self.showGrid.triggered.connect(self.show_Grid)
        self.showMGrid = QtWidgets.QAction(QtGui.QIcon('blank.png'), 'MERRA-2 Grid', self)
        self.showMGrid.setShortcut('Ctrl+M')
        self.showMGrid.setStatusTip('Show MERRA-2 Grid')
        self.showMGrid.triggered.connect(self.show_MGrid)
        self.saveView = QtWidgets.QAction(QtGui.QIcon('camera.png'), 'Save View', self)
        self.saveView.setShortcut('Ctrl+V')
        self.saveView.triggered.connect(self.save_View)
        self.showMany = QtWidgets.QAction(QtGui.QIcon('blank.png'), 'Show many Areas', self)
        self.showMany.setShortcut('Ctrl+A')
        self.showMany.setStatusTip('Show many Areas')
        self.showMany.triggered.connect(self.show_Many)
        viewMenu = menubar.addMenu('&View')
        viewMenu.addAction(self.showRuler)
        viewMenu.addAction(self.showGrid)
        viewMenu.addAction(self.showMGrid)
        viewMenu.addAction(self.showMany)
        viewMenu.addAction(self.saveView)
        self.editIni = QtWidgets.QAction(QtGui.QIcon('edit.png'), 'Edit Preferences File', self)
        self.editIni.setShortcut('Ctrl+E')
        self.editIni.setStatusTip('Edit Preferences')
        self.editIni.triggered.connect(self.editIniFile)
        self.editColour = QtWidgets.QAction(QtGui.QIcon('rainbow-icon.png'), 'Edit Colours', self)
        self.editColour.setShortcut('Ctrl+U')
        self.editColour.setStatusTip('Edit Colours')
        self.editColour.triggered.connect(self.editColours)
        self.editSect = QtWidgets.QAction(QtGui.QIcon('arrow.png'), 'Edit Section', self)
        self.editSect.setStatusTip('Edit Preferences Section')
        self.editSect.triggered.connect(self.editSects)
        editMenu = menubar.addMenu('P&references')
        editMenu.addAction(self.editIni)
        editMenu.addAction(self.editColour)
        editMenu.addAction(self.editSect)
        help = QtWidgets.QAction(QtGui.QIcon('help.png'), 'Help', self)
        help.setShortcut('F1')
        help.setStatusTip('Help')
        help.triggered.connect(self.showHelp)
        helpMenu = menubar.addMenu('&Help')
        helpMenu.addAction(help)
        QtWidgets.QShortcut(QtGui.QKeySequence('q'), self, self.exit)
        QtWidgets.QShortcut(QtGui.QKeySequence('x'), self, self.exit)
        self.config = configparser.RawConfigParser()
        self.config_file = 'getfiles.ini'
        self.config.read(self.config_file)
        parents = []
        try:
            parents = getParents(config.items('Parents'))
        except:
            pass
        try:
            mapc = self.config.get('Map', 'map_choice')
        except:
            mapc = ''
        try:
            mapp = self.config.get('Map', 'map' + mapc)
            for pkey, pvalue in parents:
                mapp = mapp.replace(pkey, pvalue)
            mapp = mapp.replace('$USER$', getUser())
            if not os.path.exists(mapp):
                if self.floatstatus is None:
                    self.show_FloatStatus()
                self.floatstatus.log.emit('Need to check [Map].map%s property. Resolves to %s' % (mapc, mapp))
        except:
            pass
        self.setWindowTitle('SIREN - worldwindow (' + fileVersion() + ')')
        self.setWindowIcon(QtGui.QIcon('sen_icon32.ico'))
        note = 'Map data ' + chr(169) + ' OpenStreetMap contributors CC-BY-SA ' + \
               '(http://www.openstreetmap.org/copyright)'
        self.view.statusmsg.emit(note)
        if not self.restorewindows:
            return
        screen = QtWidgets.QApplication.desktop().primaryScreen()
        scr_right = QtWidgets.QApplication.desktop().availableGeometry(screen).right()
        scr_bottom = QtWidgets.QApplication.desktop().availableGeometry(screen).bottom()
        win_width = self.sizeHint().width()
        win_height = self.sizeHint().height()
        try:
            rw = self.config.get('Windows', 'main_size').split(',')
            lst_width = int(rw[0])
            lst_height = int(rw[1])
            mp = self.config.get('Windows', 'main_pos').split(',')
            lst_left = int(mp[0])
            lst_top = int(mp[1])
            lst_right = lst_left + lst_width
            lst_bottom = lst_top + lst_height
            screen = QtWidgets.QApplication.desktop().screenNumber(QtCore.QPoint(lst_left, lst_top))
            scr_right = QtWidgets.QApplication.desktop().availableGeometry(screen).right()
            scr_left = QtWidgets.QApplication.desktop().availableGeometry(screen).left()
            if lst_right < scr_right:
                if (lst_right - win_width) >= scr_left:
                    scr_right = lst_right
                else:
                    scr_right = scr_left + win_width
            scr_bottom = QtWidgets.QApplication.desktop().availableGeometry(screen).bottom()
            scr_top = QtWidgets.QApplication.desktop().availableGeometry(screen).top()
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

    def editIniFile(self):
        dialr = EdtDialog(self.config_file)
        dialr.exec_()
        self.get_config()   # refresh config values
        comment = self.config_file + ' edited. Reload may be required.'
        self.view.statusmsg.emit(comment)

    def changeColours(self, new_color, elements):
        for el in elements:
            if el.pen().color().name() != '#000000':
                el.setPen(QtGui.QColor(new_color))
            if el.brush().color().name() != '#000000':
                el.setBrush(QtGui.QColor(new_color))

    def editColours(self):
        dialr = Colours(ini_file=self.config_file)
        dialr.exec_()
       # refresh some config values
        config = configparser.RawConfigParser()
        self.config.read(self.config_file)
        try:
            map = self.config.get('Map', 'map_choice')
        except:
            map = ''
        try:
            colours0 = self.config.items('Colors')
        except:
            pass
        if map != '':
            try:
                colours1 = self.config.items('Colors' + map)
                colours = []
                for it, col in colours0:
                    for it1, col1 in colours1:
                        if it1 == it:
                            colours.append((it1, col1))
                            break
                    else:
                         colours.append((it, col))
            except:
                colours = colours0
        else:
            colours = colours0
        for item, colour in colours:
            if item == 'ruler':
                continue
            if colour != self.view.scene().colors[item]:
                self.changeColours(colour, self.view.scene()._stationCircles[item])
        comment = 'Colours edited. Reload may be required.'
        self.view.statusmsg.emit(comment)

    def editSects(self):
        config = configparser.RawConfigParser()
        config.read(self.config_file)
        sections = sorted(config.sections())
        menu = QtWidgets.QMenu()
        stns = []
        for section in sections:
            stns.append(menu.addAction(section))
            x = self.frameGeometry().left() + 50
            y = self.frameGeometry().y() + 50
        action = menu.exec_(QtCore.QPoint(x, y))
        if action is not None:
            section = action.text()
            EditSect(section, None, ini_file=self.config_file)
            comment = section + ' Section edited. Reload may be required.'
            self.view.statusmsg.emit(comment)

    def mapToLonLat(self, p):
        p = self.mapToScene(p)
        return self.scene().mapToLonLat(p)

    def mapFromLonLat(self, p):
        p = self.view.scene().mapFromLonLat(p)
        return p

    def center(self):
        frameGm = self.frameGeometry()
        screen = QtWidgets.QApplication.desktop().screenNumber(QtWidgets.QApplication.desktop().cursor().pos())
        centerPoint = QtWidgets.QApplication.desktop().availableGeometry(screen).center()
        frameGm.moveCenter(centerPoint)
        self.move(frameGm.topLeft())

    def showHelp(self):
        dialog = displayobject.AnObject(QtWidgets.QDialog(), self.help,
                 title='Help for worldwindow (' + fileVersion() + ')', section='worldwindow')
        dialog.exec_()

    def setStatusText(self, text):
        if text == self.statusBar().currentMessage():
            return
        self.statusBar().showMessage(text)

    def save_View(self):
        outputimg = QtGui.QPixmap(self.view.width(), self.view.height())
        painter = QtGui.QPainter(outputimg)
        targetrect = QtCore.QRectF(0, 0, self.view.width(), self.view.height())
        sourcerect = QtCore.QRect(0, 0, self.view.width(), self.view.height())
        self.view.render(painter, targetrect, sourcerect)
        fname = 'worldwindow_view.png'
        fname = QtWidgets.QFileDialog.getSaveFileName(self, 'Save image file',
                fname, 'Image Files (*.png *.jpg *.bmp)')[0]
        if fname != '':
            i = fname.rfind('.')
            if i < 0:
                fname = fname + '.png'
                i = fname.rfind('.')
            outputimg.save(fname, fname[i + 1:])
            try:
                comment = 'View saved to ' + fname[fname.rfind('/') + 1:]
            except:
                comment = 'View saved to ' + fname
            self.view.statusmsg.emit(comment)
        painter.end()

    def popup(self, pos):
        where = self.view.mapToLonLat(pos)
        menu = QtWidgets.QMenu()
        act1 = 'Center here'
        ctrAction = menu.addAction(QtGui.QIcon('zoom.png'), act1)
        ctrAction.setIconVisibleInMenu(True)
        act2 = 'Position ruler here'
        rulAction = menu.addAction(QtGui.QIcon('ruler.png'), act2)
        rulAction.setIconVisibleInMenu(True)
        action = menu.exec_(self.mapToGlobal(pos))
        if action is None:
            return
        try:
            if action == noAction or action == notAction:
                return
        except:
            pass
        if action == rulAction:
            if self.view.scene().show_ruler:
                self.view.hide_Ruler()
            self.view.show_Ruler(self.view.scene().ruler, self.view.scene().ruler_ticks, pos)
            self.showRuler.setIcon(QtGui.QIcon('check-mark.png'))
            self.view.scene().show_ruler = True
            self.view.statusmsg.emit('Scale Ruler Toggled On')
        elif action == ctrAction:
            go_to = self.mapFromLonLat(where)
            self.view.centerOn(go_to)

    def show_Grid(self):
        comment = 'Grid Toggled '
        if self.view.scene().show_grid:
            self.view.scene().show_grid = False
            self.view.scene()._gridGroup.setVisible(False)
            self.showGrid.setIcon(QtGui.QIcon('blank.png'))
            comment += 'Off'
        else:
            if self.view.scene()._gridGroup is None:
                self.view.scene()._gridGroup = QtWidgets.QGraphicsItemGroup()
                self.view.scene().addItem(self.view.scene()._gridGroup)
                color = QtGui.QColor()
                color.setNamedColor((self.view.scene().colors['grid']))
                pen = QtGui.QPen(color, self.view.scene().line_width)
                pen.setJoinStyle(QtCore.Qt.RoundJoin)
                pen.setCapStyle(QtCore.Qt.RoundCap)
                for lat in range(90, -90, -10):
                    if lat == 90:
                        continue
                    fromm = self.mapFromLonLat(QtCore.QPointF(-180, lat))
                    too = self.mapFromLonLat(QtCore.QPointF(180, lat))
                    item = QtWidgets.QGraphicsLineItem(fromm.x(), fromm.y(), too.x(), too.y())
                    item.setPen(pen)
                    item.setZValue(3)
                    self.view.scene()._gridGroup.addToGroup(item)
                for lon in range(-180, 181, 10):
                    fromm = self.mapFromLonLat(QtCore.QPointF(lon, self.view.scene().map_upper_left[0]))
                    too = self.mapFromLonLat(QtCore.QPointF(lon, self.view.scene().map_lower_right[0]))
                    item = QtWidgets.QGraphicsLineItem(fromm.x(), fromm.y(), too.x(), too.y())
                    item.setPen(pen)
                    item.setZValue(3)
                    self.view.scene()._gridGroup.addToGroup(item)
            self.view.scene().show_grid = True
            self.view.scene()._gridGroup.setVisible(True)
            self.showGrid.setIcon(QtGui.QIcon('check-mark.png'))
            comment += 'On'
        self.view.statusmsg.emit(comment)

    def show_MGrid(self):
        comment = 'MERRA-2 Grid Toggled '
        if self.view.scene().show_mgrid:
            self.view.scene().show_mgrid = False
            self.view.scene()._mgridGroup.setVisible(False)
            self.showMGrid.setIcon(QtGui.QIcon('blank.png'))
            comment += 'Off'
        else:
            if self.view.scene()._mgridGroup is None:
                self.view.scene()._mgridGroup = QtWidgets.QGraphicsItemGroup()
                self.view.scene().addItem(self.view.scene()._mgridGroup)
                color = QtGui.QColor()
                color.setNamedColor((self.view.scene().colors['mgrid']))
                pen = QtGui.QPen(color, self.view.scene().line_width)
                pen.setJoinStyle(QtCore.Qt.RoundJoin)
                pen.setCapStyle(QtCore.Qt.RoundCap)
                pen2 = QtGui.QPen(color, self.view.scene().line_width)
                pen2.setJoinStyle(QtCore.Qt.RoundJoin)
           #     pen2.setStyle(QtCore.Qt.DotLine)
                pen2.setCapStyle(QtCore.Qt.RoundCap)
                lat = 85
                while lat > -85:
                    fromm = self.mapFromLonLat(QtCore.QPointF(-180, lat))
                    too = self.mapFromLonLat(QtCore.QPointF(180, lat))
                    item = QtWidgets.QGraphicsLineItem(fromm.x(), fromm.y(), too.x(), too.y())
                    if lat % 10 == 0:
                        item.setPen(pen)
                    else:
                        item.setPen(pen2)
                    item.setZValue(3)
                    self.view.scene()._mgridGroup.addToGroup(item)
                    lat -= .5
                lon = -180
                while lon < 180:
                    fromm = self.mapFromLonLat(QtCore.QPointF(lon, self.view.scene().map_upper_left[0]))
                    too = self.mapFromLonLat(QtCore.QPointF(lon, self.view.scene().map_lower_right[0]))
                    item = QtWidgets.QGraphicsLineItem(fromm.x(), fromm.y(), too.x(), too.y())
                    if lon % 10 == 0:
                        item.setPen(pen)
                    else:
                        item.setPen(pen2)
                    item.setZValue(3)
                    self.view.scene()._mgridGroup.addToGroup(item)
                    lon += 0.625
            self.view.scene().show_mgrid = True
            self.view.scene()._mgridGroup.setVisible(True)
            self.showMGrid.setIcon(QtGui.QIcon('check-mark.png'))
            comment += 'On'
        self.view.statusmsg.emit(comment)

    def show_Ruler(self):
        comment = 'Scale Ruler Toggled '
        if self.view.scene().show_ruler:
            self.view.hide_Ruler()
            self.showRuler.setIcon(QtGui.QIcon('blank.png'))
            self.view.scene().show_ruler = False
            comment += 'Off'
        else:
            self.view.show_Ruler(self.view.scene().ruler, self.view.scene().ruler_ticks)
            self.showRuler.setIcon(QtGui.QIcon('check-mark.png'))
            self.view.scene().show_ruler = True
            comment += 'On'
        self.view.statusmsg.emit(comment)

    def show_Many(self):
        self.view.drawMany()

    def exit(self):
        self.close()

    def closeEvent(self, event):
        if self.restorewindows:
            updates = {}
            lines = []
            add = int((self.frameSize().width() - self.size().width()) / 2)  # need to account for border
            lines.append('main_pos=%s,%s' % (str(self.pos().x() + add), str(self.pos().y() + add)))
            lines.append('main_size=%s,%s' % (str(self.size().width()), str(self.size().height())))
            lines.append('main_view=%s,%s,%s,%s' % (str(self.view.mapToScene(0, 0).x()), str(self.view.mapToScene(0, 0).y()),
                         str(self.view.mapToScene(self.view.width(), self.view.height()).x()),
                         str(self.view.mapToScene(self.view.width(), self.view.height()).y())))
            updates['Windows'] = lines
            SaveIni(updates, ini_file=self.config_file)
        self.view.tellarea.emit(['goodbye'], '')


if '__main__' == __name__:
    app = QtWidgets.QApplication(sys.argv)
    scene = WorldScene()
    openNewWindow = WorldWindow(None, scene)
    openNewWindow.show()
    app.exec_()
    app.deleteLater()
    sys.exit()
