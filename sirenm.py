#!/usr/bin/python
#
#  Copyright (C) 2015-2016 Sustainable Energy Now Inc., Angus King
#
#  siren.py - This file is part of SIREN.
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

import csv
import math
import os
import sys
# import types
import matplotlib
matplotlib.use('TkAgg')
import webbrowser
import xlrd
import xlwt
from functools import partial

import ConfigParser  # decode .ini file
from PyQt4 import QtCore, QtGui
from PyQt4.QtCore import SIGNAL
import subprocess

from colours import Colours
import displaytable
import displayobject
import newstation
from plotweather import PlotWeather
from powermodel import PowerModel
from senuser import getUser
from station import Station, Stations
from wascene import WAScene
from editini import EdtDialog, EditTech, EditSect, SaveIni
from dijkstra_4 import Shortest
from credits import Credits, fileVersion
from viewresource import Resource
from floatmenu import FloatMenu
from floatlegend import FloatLegend
from floatstatus import FloatStatus
from sirenicons import Icons


def p2str(p):
    return '(%.4f,%.4f)' % (p.y(), p.x())


class Description(QtGui.QDialog):
    def __init__(self, who, desc='', parent=None):
        super(Description, self).__init__(parent)
        layout = QtGui.QVBoxLayout(self)
        layout.addWidget(QtGui.QLabel('Set Description for ' + who))
        self.text = QtGui.QPlainTextEdit()
        self.text.setPlainText(desc)
        layout.addWidget(self.text)
         # OK and Cancel buttons
        buttons = QtGui.QDialogButtonBox(
            QtGui.QDialogButtonBox.Ok | QtGui.QDialogButtonBox.Cancel,
            QtCore.Qt.Horizontal, self)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.setWindowTitle('SIREN - Save Scenario')

    def description(self):
        return str(self.text.toPlainText())

     # static method to create the dialog and return
    @staticmethod
    def getDescription(who, desc='', parent=None):
        dialog = Description(who, desc, parent)
        result = dialog.exec_()
        return (dialog.description())  #, result == QtGui.QDialog.Accepted)


class MapView(QtGui.QGraphicsView):
    def __init__(self, scene, zoom=.8):
        QtGui.QGraphicsView.__init__(self, scene)
        self.zoom = zoom
        QtGui.QShortcut(QtGui.QKeySequence('pgup'), self, self.zoomIn)
        QtGui.QShortcut(QtGui.QKeySequence('pgdown'), self, self.zoomOut)
        QtGui.QShortcut(QtGui.QKeySequence('Ctrl++'), self, self.zoomIn)
        QtGui.QShortcut(QtGui.QKeySequence('Ctrl+-'), self, self.zoomOut)
    #     self.scene = scene
        self.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)

        self._drag_start = None
        self.setMouseTracking(True)

        self._shown = set()
        self._show_places = set(xrange(1, 10))
        self._station_to_move = None
        self._move_station = False
        self._move_grid = False
        self._grid_start = None
        self._grid_line = None
        self._grid_coord = None
        self.ruler_items = []
        self.legend_items = []
        self.legend_pos = None
        self.trace_items = []
        self.resource_items = []
        self.resource_range = None

#        QtGui.QShortcut(QtGui.QKeySequence('t'), self, self.toggleTotal)
    def destinationxy(self, lon1, lat1, bearing, distance):
        """
        Given a start point, initial bearing, and distance, calculate
        the destination point and final bearing travelling along a
        (shortest distance) great circle arc
        """
        radius = 6367.  # km is the radius of the Earth
     # convert decimal degrees to radians
        ln1, lt1, baring = map(math.radians, [lon1, lat1, bearing])
     # "reverse" haversine formula
        lat2 = math.asin(math.sin(lt1) * math.cos(distance / radius) +
               math.cos(lt1) * math.sin(distance / radius) * math.cos(baring))
        lon2 = ln1 + math.atan2(math.sin(baring) * math.sin(distance / radius) * math.cos(lt1),
               math.cos(distance / radius) - math.sin(lt1) * math.sin(lat2))
        return QtCore.QPointF(math.degrees(lon2), math.degrees(lat2))

    def mapToLonLat(self, p):
        p = self.mapToScene(p)
        return self.scene().mapToLonLat(p)

    def mapFromLonLat(self, p):
        p = self.scene().mapFromLonLat(p)
        return p

    def wheelEvent(self, event):
        delta = -event.delta() / (100 * (1 + 1 - self.zoom))
        zoom = math.pow(self.zoom, delta)
        self.scale(zoom, zoom)

    def zoomIn(self):
        self.scale(1 / self.zoom, 1 / self.zoom)

    def zoomOut(self):
        self.scale(self.zoom, self.zoom)

    def delStation(self, st):  # remove stations graphic items
        for itm in self.scene()._stationGroups[st.name]:
            self.scene().removeItem(itm)
        del self.scene()._stationGroups[st.name]
        for i in range(len(self.scene().lines.lines) - 1, -1, -1):
            if self.scene().lines.lines[i].name == st.name:
                del self.scene().lines.lines[i]

    def mousePressEvent(self, event):
        if QtCore.Qt.LeftButton == event.button():
            where = self.mapToLonLat(event.pos())
            if self._move_station:
                self.delStation(self._station_to_move)
                self._station_to_move.lon = where.x()
                self._station_to_move.lat = where.y()
                self.scene().addStation(self._station_to_move)
                comment = 'Placed station %s (%s MW at %s,%s) ' % (self._station_to_move.name,
                          '{:0.0f}'.format(self._station_to_move.capacity),
                          '{:0.4f}'.format(self._station_to_move.lat),
                          '{:0.4f}'.format(self._station_to_move.lon))
                self.emit(SIGNAL('statusmsg'), comment)
                if not self.scene().show_station_name:
                    self.scene()._current_name.setText(self._station_to_move.name)
                    pp = self.mapFromLonLat(QtCore.QPointF(where.x(), where.y()))
                    self.scene()._current_name.setPos(pp + QtCore.QPointF(1.5, -0.5))
                    if self._station_to_move.technology[:6] == 'Fossil':
                        self.scene()._current_name.setBrush(QtGui.QColor(self.scene().colors['fossil_name']))
                    else:
                        self.scene()._current_name.setBrush(QtGui.QColor(self.scene().colors['station_name']))
                self._move_station = False
                self.scene().refreshGrid()
            elif self._move_grid:      # allow multi point grid line
                color = QtGui.QColor()
                color.setNamedColor(self.scene().colors['new_grid'])
                pen = QtGui.QPen(color, self.scene().line_width)
                pen.setJoinStyle(QtCore.Qt.RoundJoin)
                pen.setCapStyle(QtCore.Qt.RoundCap)
                end = QtCore.QPointF(self.mapToScene(event.pos()))
                lin = QtGui.QGraphicsLineItem(QtCore.QLineF(self._grid_start, end))
                lin.setPen(pen)
                lin.setZValue(0)
                coord = self.scene().mapToLonLat(self._grid_start)
                self.scene().addItem(lin)
                if self._grid_line is None:
                    self._grid_line = [lin]
                    self._grid_coord = [[coord.y(), coord.x()]]
                else:
                    self._grid_line.append(lin)
                    self._grid_coord.append([coord.y(), coord.x()])
                self._grid_start = end
            else:
                try:
                    town, to_dist = self.scene()._towns.Nearest(where.y(), where.x(), distance=True)
                    town_name = town.name
                except:
                    town_name = 'No towns found'
                    to_dist = 0
                try:
                    station, st_dist = self.scene()._stations.Nearest(where.y(), where.x(), distance=True,
                        fossil=self.scene().show_fossil)
                    if self.scene().center_on_click:
                        go_to = self.mapFromLonLat(QtCore.QPointF(station.lon, station.lat))
                        self.centerOn(go_to)
                    if not self.scene().show_station_name:
                        self.scene()._current_name.setText(station.name)
                        pp = self.mapFromLonLat(QtCore.QPointF(station.lon, station.lat))
                        self.scene()._current_name.setPos(pp + QtCore.QPointF(1.5, -0.5))
                        if station.technology[:6] == 'Fossil':
                            self.scene()._current_name.setBrush(QtGui.QColor(self.scene().colors['fossil_name']))
                        else:
                            self.scene()._current_name.setBrush(QtGui.QColor(self.scene().colors['station_name']))
   # highlight grid line
                    self.emit(SIGNAL('statusmsg'), p2str(self.mapToLonLat(event.pos())) + ' ' +
                      station.name + ' ({:0.0f}'.format(station.capacity) + ' MW; {:0.0f} Km away; '.format(st_dist) +
                      'Nearest town: ' + town_name + ' {:0.0f} Km away)'.format(to_dist))
                except:
                    self.emit(SIGNAL('statusmsg'), p2str(self.mapToLonLat(event.pos())) + ' ' +
                      'Nearest town: ' + town_name + ' {:0.0f} Km away)'.format(to_dist))
                self._drag_start = QtCore.QPoint(event.pos())
                hb = self.horizontalScrollBar()
                vb = self.verticalScrollBar()
                self._sb_start = QtCore.QPoint(hb.value(), vb.value())

    def mouseMoveEvent(self, event):
        if QtCore.Qt.LeftButton == event.buttons() and self._drag_start:
            delta = event.pos() - self._drag_start
            val = self._sb_start - delta
            hb = self.horizontalScrollBar().setValue(val.x())
            vb = self.verticalScrollBar().setValue(val.y())
        else:
            p = self.mapToScene(event.pos())
            x, y = map(int, [p.x(), p.y()])
            try:
                stations = self.view.scene().positions()[(x, y)]
                for st in self._shown.difference(stations):
                    self._show_places.add(st.hideInfo())
                for st in set(stations).difference(self._shown):
                    start = self.mapToGlobal(self.pos())
                    start.setX(start.x() + self.width())
                    place = sorted(self._show_places)[0]
                    st.showInfo(place, start)
                    self._show_places.remove(place)
                self._shown = set(stations)
                return
            except KeyError:
                for st in self._shown:
                    self._show_places.add(st.hideInfo())
                self._shown = set()
            except:
                pass
        pl = self.mapToLonLat(event.pos())
      #   self.emit(SIGNAL('statusmsg'), p2str(pl) + ' ' + p2str(event.pos()))

    def mouseDoubleClickEvent(self, event):
        if self._move_grid:
            p = self.mapToLonLat(event.pos())
            self._grid_coord.append([p.y(), p.x()])
            self.delStation(self._station_to_move)
            self._station_to_move.grid_line = '(' + str(self._grid_coord[0][0]) + ',' + str(self._grid_coord[0][1]) + ')'
            for i in range(1, len(self._grid_coord)):
                self._station_to_move.grid_line += ',(' + str(self._grid_coord[i][0]) + ',' + str(self._grid_coord[i][1]) + ')'
            self.scene().addStation(self._station_to_move)
            if not self.scene().show_station_name:
                self.scene()._current_name.setText(self._station_to_move.name)
                pp = self.mapFromLonLat(QtCore.QPointF(self._station_to_move.lon, self._station_to_move.lat))
                self.scene()._current_name.setPos(pp + QtCore.QPointF(1.5, -0.5))
                if self._station_to_move.technology[:6] == 'Fossil':
                    self.scene()._current_name.setBrush(QtGui.QColor(self.scene().colors['fossil_name']))
                else:
                    self.scene()._current_name.setBrush(QtGui.QColor(self.scene().colors['station_name']))
            self._move_grid = False
            self._grid_start = None
            for itm in self._grid_line:
                self.scene().removeItem(itm)
            self._grid_line = None
            self._grid_coord = None
            self.scene().refreshGrid()

    def mouseReleaseEvent(self, event):
        if QtCore.Qt.LeftButton == event.button():
            if self._move_grid:
                pass
            else:
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
            self.ruler_items.append(QtGui.QGraphicsLineItem(QtCore.QLineF(start, end)))
            self.ruler_items[-1].setPen(pen)
            self.ruler_items[-1].setZValue(0)
            self.scene().addItem(self.ruler_items[-1])
            toll = self.destinationxy(frll.x(), frll.y(), 90, ruler)
            end = self.mapFromLonLat(toll)
            self.ruler_items.append(QtGui.QGraphicsLineItem(QtCore.QLineF(start, end)))
            self.ruler_items[-1].setPen(pen)
            self.ruler_items[-1].setZValue(0)
            self.scene().addItem(self.ruler_items[-1])
            for i in range(ticks + 1):
                strt = self.destinationxy(frll.x(), frll.y(), 90, ruler * i / ticks)
                start = self.mapFromLonLat(strt)
                toll = self.destinationxy(strt.x(), strt.y(), 0, ruler / 50)
                end = self.mapFromLonLat(toll)
                self.ruler_items.append(QtGui.QGraphicsLineItem(QtCore.QLineF(start, end)))
                self.ruler_items[-1].setPen(pen)
                self.ruler_items[-1].setZValue(0)
                self.scene().addItem(self.ruler_items[-1])
            self.ruler_items.append(QtGui.QGraphicsSimpleTextItem(str(int(ruler)) + ' Km'))
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
                toll = self.destinationxy(strt.x(), strt.y(), 90, ruler / 50)
                end = self.mapFromLonLat(toll)
                self.ruler_items.append(QtGui.QGraphicsLineItem(QtCore.QLineF(start, end)))
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

    def show_Legend(self, pos=None, where=None):
        self.legend_pos = pos
        if where is not None:
            p = self.mapFromLonLat(QtCore.QPointF(where.x(), where.y()))
            p = QtCore.QPoint(p.x(), p.y())
        tech_sizes = {}
        tot_capacity = 0.
        tot_stns = 0
        for st in self.scene()._stations.stations:
            if st.technology[:6] == 'Fossil' and not self.scene().show_fossil:
                continue
            if st.technology not in tech_sizes:
                tech_sizes[st.technology] = 0.
        #     if st.technology[:6] == 'Fossil' or st.technology == 'Rooftop PV':
         #       continue
            if st.technology == 'Wind':
                tech_sizes[st.technology] += self.scene().areas[st.technology] * float(st.no_turbines) \
                                             * pow((st.rotor * .001), 2)
            else:
                tech_sizes[st.technology] += self.scene().areas[st.technology] * float(st.capacity)
            tot_capacity += st.capacity
            tot_stns += 1
        self.legend_items = []
        if pos is None and where is None:
            itm = QtGui.QGraphicsSimpleTextItem('Legend')
            new_font = itm.font()
            new_font.setPointSizeF(self.scene().width() / 90.)
            fh = int(QtGui.QFontMetrics(new_font).height() * 1.1)
            p = QtCore.QPointF(self.scene().lower_right[0] + fh, self.scene().upper_left[1] + fh)
            frll = self.scene().mapToLonLat(p)
            p = self.mapFromLonLat(QtCore.QPointF(frll.x(), frll.y()))
            p = QtCore.QPoint(p.x(), p.y())
        elif where is None:
            p = self.mapToScene(pos)
            p = QtCore.QPoint(p.x(), p.y())
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
        self.legend_items.append(QtGui.QGraphicsSimpleTextItem('Legend'))
        new_font = self.legend_items[-1].font()
        new_font.setPointSizeF(self.scene().width() / 90.)
        fh = int(QtGui.QFontMetrics(new_font).height() * 1.1)
        self.legend_items[-1].setFont(new_font)
        self.legend_items[-1].setPos(p.x(), p.y())
        self.legend_items[-1].setBrush(QtGui.QColor(color))
        self.legend_items[-1].setZValue(0)
        self.scene().addItem(self.legend_items[-1])
        p.setY(p.y() + fh)
        tot_area = 0
        for key, value in iter(sorted(tech_sizes.iteritems())):
            tot_area += value
            self.legend_items.append(QtGui.QGraphicsPolygonItem(QtGui.QPolygonF([QtCore.QPointF(p.x(), p.y() + int(fh * .45)),
                                     QtCore.QPointF(p.x() + int(fh * .4), p.y() + int(fh * .05)),
                                     QtCore.QPointF(p.x() + int(fh * .8), p.y() + int(fh * .45)),
                                     QtCore.QPointF(p.x() + int(fh * .4), p.y() + int(fh * .85)),
                                     QtCore.QPointF(p.x(), p.y() + int(fh * .45))])))
          #   if self.scene().station_square:
          #       self.legend_items.append(QtGui.QGraphicsRectItem(p.x(), p.y(), int(fh * .8), int(fh * .8)))
          #   else:
          #       self.legend_items.append(QtGui.QGraphicsEllipseItem(p.x(), p.y(), int(fh * .8), int(fh * .8)))
            self.legend_items[-1].setBrush(QtGui.QColor(self.scene().colors[key]))
            if self.scene().colors['border'] != '':
                self.legend_items[-1].setPen(QtGui.QColor(self.scene().colors['border']))
            else:
                self.legend_items[-1].setPen(QtGui.QColor(self.scene().colors[key]))
            self.legend_items[-1].setZValue(1)
            self.scene().addItem(self.legend_items[-1])
            self.legend_items.append(QtGui.QGraphicsSimpleTextItem(key))
            self.legend_items[-1].setFont(new_font)
            self.legend_items[-1].setPos(p.x() + fh, p.y())
            self.legend_items[-1].setBrush(QtGui.QColor(color))
            self.legend_items[-1].setZValue(0)
            self.scene().addItem(self.legend_items[-1])
            p.setY(p.y() + fh)
        if self.scene().show_capacity or self.scene().show_generation:
            txt = 'Circles show relative\n  '
            if self.scene().show_capacity:
                txt += 'capacity in MW'
            else:
                txt += 'generation in MWh'
            self.legend_items.append(QtGui.QGraphicsSimpleTextItem(txt))
            new_font = self.legend_items[-1].font()
            new_font.setPointSizeF(self.scene().width() / 110.)
            fh = int(QtGui.QFontMetrics(new_font).height() * 1.1)
            self.legend_items[-1].setFont(new_font)
            self.legend_items[-1].setPos(p.x(), p.y())
            self.legend_items[-1].setBrush(QtGui.QColor(color))
            self.legend_items[-1].setZValue(0)
            self.scene().addItem(self.legend_items[-1])
            p.setY(p.y() + fh * 2)
            if self.scene().show_capacity:
                if tot_stns > 0:
                    avg_capacity = tot_capacity / tot_stns
                    last_yd = 0
                    while True:
                        avg_capacity = round(avg_capacity, -(len(str(int(avg_capacity))) - 1))
                        size = math.sqrt(avg_capacity * self.scene().capacity_area / math.pi)
                        frl = self.scene().mapToLonLat(p)
                        e = self.destinationxy(frl.x(), frl.y(), 90., size)
                        pe = self.mapFromLonLat(QtCore.QPointF(e.x(), frl.y()))
                        n = self.destinationxy(frl.x(), frl.y(), 0., size)
                        pn = self.mapFromLonLat(QtCore.QPointF(frl.x(), n.y()))
                        x_d = p.x() - pe.x()
                        y_d = pn.y() - p.y()
                        if abs(y_d) <= fh * 4 or y_d == last_yd:
                            break
                        last_yd = y_d
                        avg_capacity = avg_capacity / 2
                    el = QtGui.QGraphicsEllipseItem(p.x() - x_d, p.y() - y_d, x_d, y_d)
                    self.legend_items.append(el)
                    color.setAlphaF(self.scene().capacity_opacity)
                    self.legend_items[-1].setBrush(color)
                    color.setAlphaF(1)
                    self.legend_items[-1].setPen(color)
                    self.scene().addItem(self.legend_items[-1])
                    txt = ' = %s MW' % str(int(avg_capacity))
                    self.legend_items.append(QtGui.QGraphicsSimpleTextItem(txt))
                    self.legend_items[-1].setFont(new_font)
                    self.legend_items[-1].setPos(pe.x() + 5, p.y() - y_d / 2 - fh / 2)
                    self.legend_items[-1].setBrush(QtGui.QColor(color))
                    self.legend_items[-1].setZValue(0)
                    self.scene().addItem(self.legend_items[-1])
                    y_d = max(abs(y_d), fh)
                    p.setY(p.y() + y_d)
        txt = 'Station '
        if self.scene().station_square:
            txt += 'Squares'
        else:
            if self.scene().show_capacity or self.scene().show_generation:
                txt += '(Inner) '
            txt += 'Circles'
        txt += ' show\n estimated area in sq. Km'
        self.legend_items.append(QtGui.QGraphicsSimpleTextItem(txt))
        new_font = self.legend_items[-1].font()
        new_font.setPointSizeF(self.scene().width() / 110.)
        fh = int(QtGui.QFontMetrics(new_font).height() * 1.1)
        self.legend_items[-1].setFont(new_font)
        self.legend_items[-1].setPos(p.x(), p.y())
        self.legend_items[-1].setBrush(QtGui.QColor(color))
        self.legend_items[-1].setZValue(0)
        self.scene().addItem(self.legend_items[-1])
        p.setY(p.y() + fh * 2)
        self.legend_items.append(QtGui.QGraphicsSimpleTextItem('Total land area by\n  technology (' +
          '{:0.0f}'.format(tot_area) + ' sq. Km):'))
        self.legend_items[-1].setFont(new_font)
        self.legend_items[-1].setPos(p.x(), p.y())
        self.legend_items[-1].setBrush(QtGui.QColor(color))
        self.legend_items[-1].setZValue(0)
        self.scene().addItem(self.legend_items[-1])
        p.setY(p.y() + fh * 2)
        frl = self.scene().mapToLonLat(p)
        for key, value in sorted(tech_sizes.iteritems(), key=lambda x: x[1]):
            size = math.sqrt(value)
            e = self.destinationxy(frl.x(), frl.y(), 90., size)
            s = self.destinationxy(frl.x(), frl.y(), 180., size)
            p = self.mapFromLonLat(QtCore.QPointF(frl.x(), frl.y()))
            pe = self.mapFromLonLat(QtCore.QPointF(e.x(), frl.y()))
            ps = self.mapFromLonLat(QtCore.QPointF(frl.x(), s.y()))
            se = self.mapFromLonLat(QtCore.QPointF(e.x(), s.y()))
            x_d = pe.x() - p.x()
            y_d = ps.y() - p.y()
            self.legend_items.append(QtGui.QGraphicsRectItem(p.x(), p.y(), x_d, y_d))
            self.legend_items[-1].setBrush(QtGui.QColor(self.scene().colors[key]))
            if self.scene().colors['border'] != '':
                self.legend_items[-1].setPen(QtGui.QColor(self.scene().colors['border']))
            else:
                self.legend_items[-1].setPen(QtGui.QColor(self.scene().colors[st.technology]))
            self.legend_items[-1].setZValue(1)
            self.scene().addItem(self.legend_items[-1])
            frl.setY(s.y())
        x = self.scene().upper_left[0]
        y = self.scene().upper_left[1]
        w = self.scene().lower_right[0]
        h = self.scene().lower_right[1]
        self.scene().setSceneRect(-w * 0.05, -h * 0.05, w * 1.1, h * 1.1)
        x = y = r = b = 0
        for itm in self.legend_items:
            if isinstance(itm, QtGui.QGraphicsSimpleTextItem):
                xi = itm.pos().x()
                yi = itm.pos().y()
            else:
                xi = itm.boundingRect().x()
                yi = itm.boundingRect().y()
            x = min(x, xi)
            y = min(y, yi)
            r = max(r, xi + itm.boundingRect().width())
            b = max(b, yi + itm.boundingRect().height())
        if x > self.scene().upper_left[0]:
            x = self.scene().upper_left[0]
        if y > self.scene().upper_left[1]:
            y = self.scene().upper_left[1]
        if r < self.scene().lower_right[0]:
            r = self.scene().lower_right[0]
        if b < self.scene().lower_right[1]:
            b = self.scene().lower_right[1]
        w = r - x
        h = b - y
        x = round(x - w * 0.05, 1)
        y = round(y - h * 0.05, 1)
        r = round(r + w * 0.1, 1)
        b = round(b + h * 0.1, 1)
        self.scene().setSceneRect(x, y, r, b)

    def hide_Legend(self, pos=None):
        for i in range(len(self.legend_items)):
            self.scene().removeItem(self.legend_items[i])
        del self.legend_items
        if pos is None:
            self.legend_pos = None
        x = self.scene().upper_left[0]
        y = self.scene().upper_left[1]
        w = self.scene().lower_right[0]
        h = self.scene().lower_right[1]
        self.scene().setSceneRect(-w * 0.05, -h * 0.05, w * 1.1, h * 1.1)

    def clear_Trace(self):
        try:
            for i in range(len(self.trace_items)):
                self.scene().removeItem(self.trace_items[i])
            del self.trace_items
        except:
            pass

    def traceGrid(self, station, coords=None):
        self.clear_Trace()
        self.trace_items = []
        color = QtGui.QColor()
        color.setNamedColor((self.scene().colors['grid_trace']))
        pen = QtGui.QPen(color, self.scene().line_width)
        pen.setJoinStyle(QtCore.Qt.RoundJoin)
        pen.setCapStyle(QtCore.Qt.RoundCap)
        if coords is None:
            for li in range(len(self.scene().lines.lines)):
                if self.scene().lines.lines[li].name == station.name:
                    break
            else:
                return
            nearest = 99999
            j = -1
            dims = self.scene().lines.lines[li].coordinates[0]
            for i in range(len(self.scene().load_centre)):
                thisone = self.scene().lines.actualDistance(self.scene().load_centre[i][1],
                          self.scene().load_centre[i][2], dims[0], dims[1])
                if thisone < nearest:
                    nearest = thisone
                    j = i
            path = Shortest(self.scene().lines.lines, self.scene().lines.lines[li].coordinates[0],
                   [self.scene().load_centre[j][1], self.scene().load_centre[j][2]], self.scene().grid_lines)
            route = path.getPath()
#           check we don't go through another load_centre
            if len(self.scene().load_centre) > 1:
                for co in range(len(route) - 1, 0, -1):
                    for i in range(len(self.scene().load_centre)):
                        if route[co][0] == self.scene().load_centre[i][1] and \
                          route[co][1] == self.scene().load_centre[i][2]:
                            route = route[i:]
                            break
        else:
            route = coords
        grid_path_len = 0.
        st_scn = self.mapFromLonLat(QtCore.QPointF(route[0][1], route[0][0]))
        for i in range(1, len(route)):
            en_scn = self.mapFromLonLat(QtCore.QPointF(route[i][1], route[i][0]))
            grid_path_len += self.scene().lines.actualDistance(route[i - 1][0], route[i - 1][1], route[i][0], route[i][1])
            self.trace_items.append(QtGui.QGraphicsLineItem(QtCore.QLineF(st_scn, en_scn)))
            self.trace_items[-1].setPen(pen)
            self.trace_items[-1].setZValue(3)
            self.scene().addItem(self.trace_items[-1])
            st_scn = en_scn
        return grid_path_len

    def clear_Resource(self):
        try:
            for i in range(len(self.resource_items)):
                self.scene().removeItem(self.resource_items[i])
            del self.resource_items
        except:
            pass

    def resourceGrid(self, option, resource_window=None, year=None):
        def in_map(a_min, a_max, b_min, b_max):
            return (a_min <= b_max) and (b_min <= a_max)
        if option == 'hide':
            self.clear_Resource()
            return
        if self.scene().resource_grid == '':
            return
        if year is None:
            r_year = self.scene().base_year
        else:
            r_year = year
        period, variable, steps, opacity, colours = resource_window.Results()
        i = period.find('_')
        if i > 0:
            resource_file = self.scene().resource_grid.replace('$YEAR$', period[:i])
        else:
            resource_file = self.scene().resource_grid.replace('$YEAR$', period[:4])
        if os.path.exists(resource_file):
            self.clear_Resource()
            self.resource_items = []
            lo_valu = 99999.
            hi_valu = 0.
            calc_minmax = True
            cells = []
            if resource_file[-4:] == '.xls' or resource_file[-5:] == '.xlsx':
                var = {}
                workbook = xlrd.open_workbook(resource_file)
                worksheet = workbook.sheet_by_index(0)
                num_rows = worksheet.nrows - 1
                num_cols = worksheet.ncols - 1
#               get column names
                curr_col = -1
                while curr_col < num_cols:
                    curr_col += 1
                    var[worksheet.cell_value(0, curr_col)] = curr_col
                curr_row = 1
                if worksheet.cell_value(curr_row, var['Period']) == 'Min.' and \
                  worksheet.cell_value(curr_row + 1, var['Period']) == 'Max.':
                    calc_minmax = False
                    lo_valu = worksheet.cell_value(curr_row, var[variable])
                    hi_valu = worksheet.cell_value(curr_row + 1, var[variable])
                    curr_row += 1
                while curr_row < num_rows:
                    curr_row += 1
                    a_lat = float(worksheet.cell_value(curr_row, var['Latitude']))
                    a_lon = float(worksheet.cell_value(curr_row, var['Longitude']))
                    if in_map(a_lat - 0.25, a_lat + 0.25, self.scene().map_lower_right[0],
                              self.scene().map_upper_left[0]) \
                      and in_map(a_lon - 0.3333, a_lon + 0.3333,
                              self.scene().map_upper_left[1], self.scene().map_lower_right[1]):
                        try:
                            if worksheet.cell_value(curr_row, var['Period']) == period:
                                cells.append([float(worksheet.cell_value(curr_row, var['Latitude'])),
                                             float(worksheet.cell_value(curr_row, var['Longitude'])),
                                             worksheet.cell_value(curr_row, var[variable])])
                            if calc_minmax:
                                if worksheet.cell_value(curr_row, var[variable]) < lo_valu:
                                    lo_valu = worksheet.cell_value(curr_row, var[variable])
                                if worksheet.cell_value(curr_row, var[variable]) > hi_valu:
                                    hi_valu = worksheet.cell_value(curr_row, var[variable])
                        except:
                            pass
            else:
                resource = open(resource_file)
                things = csv.DictReader(resource)
                for cell in things:
                    a_lat = float(cell['Latitude'])
                    a_lon = float(cell['Longitude'])
                    if in_map(a_lat - 0.25, a_lat + 0.25, self.scene().map_lower_right[0],
                              self.scene().map_upper_left[0]) \
                      and in_map(a_lon - 0.3333, a_lon + 0.3333,
                              self.scene().map_upper_left[1], self.scene().map_lower_right[1]):
                        if cell['Period'] == period:
                            cells.append(cell['Latitude'], cell['Longitude'], cell[variable])
                        if cell[variable] < lo_valu:
                            lo_valu = cell[variable]
                        if cell[variable] > hi_valu:
                            hi_valu = cell[variable]
                resource.close()
        else:
            return
        if steps > 0:
            incr = (hi_valu - lo_valu) / steps
        else:
            lo_colour = []
            hi_colour = []
            for i in range(3):
                lo_1 = int(colours[0][i * 2 + 1:i * 2 + 3], base=16)
                lo_colour.append(lo_1)
                hi_1 = int(colours[-1][i * 2 + 1:i * 2 + 3], base=16)
                hi_colour.append(hi_1)
        lo_per = 99999.
        hi_per = 0.
        lons = []
        lon_cell = .3125
        for cell in cells:
            if cell[1] not in lons:
                lons.append(cell[1])
        if len(lons) > 1:
            lons = sorted(lons)
            lon_cell = (lons[1] - lons[0]) / 2.
            del lons
        for cell in cells:
            p = self.mapFromLonLat(QtCore.QPointF(cell[1] - lon_cell, cell[0] + .25))
            pe = self.mapFromLonLat(QtCore.QPointF(cell[1] + lon_cell, cell[0] + .25))
            ps = self.mapFromLonLat(QtCore.QPointF(cell[1] - lon_cell, cell[0] - .25))
            se = self.mapFromLonLat(QtCore.QPointF(cell[1] + lon_cell, cell[0] - .25))
            x_d = pe.x() - p.x()
            y_d = ps.y() - p.y()
            self.resource_items.append(QtGui.QGraphicsRectItem(p.x(), p.y(), x_d, y_d))
            if steps > 0:
                step = int(round((cell[2] - lo_valu) / incr))
                a_colour = QtGui.QColor(colours[step])
            else:
                colr = []
                pct = (cell[2] - lo_valu) / (hi_valu - lo_valu)
                for i in range(3):
                    colr.append(((hi_colour[i] - lo_colour[i]) * pct + lo_colour[i]) / 255.)
                a_colour = QtGui.QColor()
                a_colour.setRgbF(colr[0], colr[1], colr[2])
            self.resource_items[-1].setBrush(a_colour)
            self.resource_items[-1].setPen(a_colour)
            self.resource_items[-1].setOpacity(opacity)
            self.resource_items[-1].setZValue(1)
            self.scene().addItem(self.resource_items[-1])
            if cell[2] < lo_per:
                lo_per = cell[2]
            if cell[2] > hi_per:
                hi_per = cell[2]
        self.resource_range = [variable, period, lo_valu, hi_valu, lo_per, hi_per]


class MainWindow(QtGui.QMainWindow):
    mySignal = QtCore.pyqtSignal()

    def get_config(self):
        config = ConfigParser.RawConfigParser()
        if len(sys.argv) > 1:
            self.config_file = sys.argv[1]
        else:
            self.config_file = 'SIREN.ini'
        config.read(self.config_file)
        try:
            self.base_year = config.get('Base', 'year')
        except:
            self.base_year = '2012'
        try:
            self.years = []
            years = config.get('Base', 'years')
            bits = years.split(',')
            for i in range(len(bits)):
                rngs = bits[i].split('-')
                if len(rngs) > 1:
                    for j in range(int(rngs[0].strip()), int(rngs[1].strip()) + 1):
                        self.years.append(str(j))
                else:
                    self.years.append(rngs[0].strip())
        except:
            self.years = None
        self.existing = True
        try:
            existing = config.get('Base', 'existing')
            if existing.lower() in ['false', 'no', 'off']:
                self.existing = False
        except:
            pass
        parents = []
        try:
            parents = config.items('Parents')
        except:
            pass
        try:
            self.aboutfile = config.get('Files', 'about')
            for key, value in parents:
                self.aboutfile = self.aboutfile.replace(key, value)
            self.aboutfile = self.aboutfile.replace('$USER$', getUser())
            self.aboutfile = self.aboutfile.replace('$YEAR$', self.base_year)
        except:
            self.aboutfile = ''
        try:
            self.helpfile = config.get('Files', 'help')
            for key, value in parents:
                self.helpfile = self.helpfile.replace(key, value)
            self.helpfile = self.helpfile.replace('$USER$', getUser())
            self.helpfile = self.helpfile.replace('$YEAR$', self.base_year)
        except:
            self.helpfile = ''
        try:
            self.scenarios = config.get('Files', 'scenarios')
            for key, value in parents:
                self.scenarios = self.scenarios.replace(key, value)
            self.scenarios = self.scenarios.replace('$USER$', getUser())
            self.scenarios = self.scenarios.replace('$YEAR$', self.base_year)
            i = self.scenarios.rfind('/')
            self.scenarios_filter = self.scenarios[i + 1:]
            if self.scenarios_filter[-1] != '*':
                self.scenarios_filter += '*'
            self.scenarios = self.scenarios[:i + 1]
         #    if sys.platform != 'win32' and sys.platform != 'cygwin':
            matplotlib.rcParams['savefig.directory'] = self.scenarios
        except:
            self.scenarios = ''
            self.scenarios_filter = '*'
        try:
            self.scenario = config.get('Files', 'scenario')
            for key, value in parents:
                self.scenario = self.scenario.replace(key, value)
            self.scenario = self.scenario.replace('$USER$', getUser())
            self.scenario = self.scenario.replace('$YEAR$', self.base_year)
        except:
            self.scenario = ''
        str1 = self.scenarios_filter[:-1]   # a bit of mucking about to remove duplicate leading/trailing chars
        str2 = str(QtCore.QDateTime.toString(QtCore.QDateTime.currentDateTime(), 'yyyy-MM-dd_hhmm')) + '.xls'
        ndx = 0
        for i in range(len(str1)):
            if str1[-i:] == str2[:i]:
                ndx = i
        self.new_scenario = str1 + str2[ndx:]
        self.leave_help_open = False
        try:
            leave_help_open = config.get('View', 'leave_help_open')
            if leave_help_open.lower() in ['true', 'yes', 'on']:
                self.leave_help_open = True
        except:
            pass
        self.old_grid = False
        try:
            old_grid = config.get('View', 'old_grid')
            if old_grid.lower() in ['true', 'yes', 'on']:
                self.old_grid = True
        except:
            pass
        if not self.old_grid:
            try:
                old_grid = config.get('Grid', 'trace_existing')
                if old_grid.lower() in ['true', 'yes', 'on']:
                    self.old_grid = True
            except:
                pass
        self.grid_icon = 'network.png'
        self.check_icon = 'check-mark.png'
        try:
            mb = config.get('View', 'menu_background')
            if mb.lower() != 'b':
                self.grid_icon = 'network_b.png'
                self.check_icon = 'check-mark_b.png'
        except:
            pass
        self.restorewindows = False
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
        self.rainfall = False
        try:
            rain = config.get('View', 'resource_rainfall')
            if rain.lower() in ['true', 'yes', 'on']:
                self.rainfall = True
        except:
            pass
        self.zoom = 0.8
        try:
            self.zoom = float(config.get('View', 'zoom_rate'))
            if self.zoom > 1:
                self.zoom = 1 / self.zoom
            if self.zoom > 0.95:
                self.zoom = 0.95
            elif self.zoom < 0.75:
                self.zoom = 0.75
        except:
            pass

    def __init__(self, scene):
        QtGui.QMainWindow.__init__(self)
        self.get_config()
       #  self.setWindowFlags(QtCore.Qt.FramelessWindowHint)
        self.view = MapView(scene, self.zoom)
        self.view.scale(1., 1.)
        self._mv = self.view
        self.setStatusBar(QtGui.QStatusBar())
        self.connect(self.view, SIGNAL('statusmsg'), self.setStatusText)
        self.altered_stations = False

        w = QtGui.QWidget()
        lay = QtGui.QVBoxLayout(w)
        lay.addWidget(self.view)
        self.setCentralWidget(w)

        self.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self.popup)

        openFile = QtGui.QAction(QtGui.QIcon('open.png'), 'Open', self)
        openFile.setShortcut('Ctrl+O')
        openFile.setStatusTip('Open new Scenario')
        openFile.triggered.connect(self.addScenario)
        addFile = QtGui.QAction(QtGui.QIcon('plus.png'), 'Add', self)
        addFile.setShortcut('Ctrl+A')
        addFile.setStatusTip('Add Scenario')
        addFile.triggered.connect(self.addScenario)
        saveFile = QtGui.QAction(QtGui.QIcon('save.png'), 'Save', self)
        saveFile.setShortcut('Ctrl+S')
        saveFile.setStatusTip('Save Scenario(s)')
        saveFile.triggered.connect(self.saveScenario)
        savedFile = QtGui.QAction(QtGui.QIcon('save_d.png'), 'Save + Desc.', self)
        savedFile.setStatusTip('Save Scenario(s) change description')
        savedFile.triggered.connect(self.saveScenariod)
        saveasFile = QtGui.QAction(QtGui.QIcon('save_as.png'), 'Save As...', self)
        saveasFile.setStatusTip('Save New Scenario')
        saveasFile.triggered.connect(self.saveAsScenario)
        exitModel = QtGui.QAction(QtGui.QIcon('exit.png'), 'Exit', self)
        exitModel.setShortcut('Ctrl+X')
        exitModel.setStatusTip('Exit Scenario')
        exitModel.triggered.connect(self.exit)
        quitModel = QtGui.QAction(QtGui.QIcon('quit.png'), 'Quit', self)
        quitModel.setShortcut('Ctrl+Q')
        quitModel.setStatusTip('Quit Scenario')
        quitModel.triggered.connect(self.close)
        self.addExist = QtGui.QAction(QtGui.QIcon('plus.png'), 'Add Existing', self)
        self.addExist.setStatusTip('Add Existing Stations')
        self.addExist.triggered.connect(self.addExisting)
        menubar = self.menuBar()
        self.fileMenu = menubar.addMenu('&Scenario')
        self.fileMenu.addAction(openFile)
        self.fileMenu.addAction(addFile)
        self.subMenu = self.fileMenu.addMenu('Remove')
        self.subMenu2 = self.fileMenu.addMenu('Edit Descr.')
        self.fileMenu.addAction(saveFile)
        self.fileMenu.addAction(savedFile)
        self.fileMenu.addAction(saveasFile)
        self.fileMenu.addAction(exitModel)
        self.fileMenu.addAction(quitModel)
        for scen_filter, chng, desc in self.view.scene()._scenarios:
            subFile = QtGui.QAction(QtGui.QIcon('minus.png'), scen_filter, self)
            subFile.setStatusTip('Remove Scenario ' + scen_filter)
            subFile.triggered.connect(self.removeScenario)
            self.subMenu.addAction(subFile)
            if scen_filter != 'Existing':
                subFile2 = QtGui.QAction(QtGui.QIcon('edit.png'), scen_filter, self)
                subFile2.setStatusTip('Edit Description for ' + scen_filter)
                subFile2.triggered.connect(self.editDescription)
                self.subMenu2.addAction(subFile2)
        if not self.existing:
            self.fileMenu.addAction(self.addExist)
        getPower = QtGui.QAction(QtGui.QIcon('power.png'), 'Power for ' + self.base_year, self)
        getPower.setShortcut('Ctrl+P')
        getPower.setStatusTip('Get Generation for ' + self.base_year)
        getPower.triggered.connect(self.get_Power)
        listStations = QtGui.QAction(QtGui.QIcon('power-plant-icon.png'), 'List Stations', self)
        listStations.setStatusTip('List Stations')
        listStations.setShortcut('Ctrl+L')
        listStations.triggered.connect(self.list_Stations)
        if self.old_grid:
            tip = 'List Grid'
        else:
            tip = 'List New Grid'
        listGrid = QtGui.QAction(QtGui.QIcon(self.grid_icon), tip, self)
        listGrid.setStatusTip(tip)
        listGrid.setShortcut('Ctrl+W')
        listGrid.triggered.connect(self.list_Grid)
        powerMenu = menubar.addMenu('&Power')
        powerMenu.addAction(getPower)
        if self.years is not None:
            subPowerMenu = powerMenu.addMenu('&Power for year')
            for year in self.years:
                subPower = QtGui.QAction(QtGui.QIcon('power.png'), year, self)
                subPower.triggered.connect(self.get_Power)
                subPowerMenu.addAction(subPower)
        powerMenu.addAction(listStations)
        powerMenu.addAction(listGrid)
        if self.view.scene().show_capacity:
            self.showCapacity = QtGui.QAction(QtGui.QIcon(self.check_icon), 'Capacity Circles', self)
        else:
            self.showCapacity = QtGui.QAction(QtGui.QIcon('blank.png'), 'Capacity Circles', self)
        self.showCapacity.setShortcut('Ctrl+C')
        self.showCapacity.setStatusTip('Toggle Capacity Circles')
        self.showCapacity.triggered.connect(self.show_Capacity)
        if self.view.scene().show_generation:
            self.showGeneration = QtGui.QAction(QtGui.QIcon(self.check_icon), 'Generation Circles', self)
        else:
            self.showGeneration = QtGui.QAction(QtGui.QIcon('blank.png'), 'Generation Circles', self)
        self.showGeneration.setShortcut('Ctrl+K')
        self.showGeneration.setStatusTip('Toggle Generation Circles')
        self.showGeneration.triggered.connect(self.show_Generation)
        if self.view.scene().show_station_name:
            self.showName = QtGui.QAction(QtGui.QIcon(self.check_icon), 'Station Names', self)
        else:
            self.showName = QtGui.QAction(QtGui.QIcon('blank.png'), 'Station Names', self)
        self.showName.setShortcut('Ctrl+N')
        self.showName.triggered.connect(self.show_Name)
        if self.view.scene().show_fossil:
            self.showFossil = QtGui.QAction(QtGui.QIcon(self.check_icon), 'Fossil-fueled Stations', self)
        else:
            self.showFossil = QtGui.QAction(QtGui.QIcon('blank.png'), 'Fossil-fueled Stations', self)
        self.showFossil.setShortcut('Ctrl+F')
        self.showFossil.setStatusTip('Show Fossil Stations')
        self.showFossil.triggered.connect(self.show_Fossil)
        if self.view.scene().show_ruler:
            self.showRuler = QtGui.QAction(QtGui.QIcon(self.check_icon), 'Scale Ruler', self)
        else:
            self.showRuler = QtGui.QAction(QtGui.QIcon('blank.png'), 'Scale Ruler', self)
        self.showRuler.setShortcut('Ctrl+R')
        self.showRuler.setStatusTip('Show Scale Ruler')
        self.showRuler.triggered.connect(self.show_Ruler)
        if self.view.scene().show_legend:
            self.showLegend = QtGui.QAction(QtGui.QIcon(self.check_icon), 'Show Legend', self)
        else:
            self.showLegend = QtGui.QAction(QtGui.QIcon('blank.png'), 'Show Legend', self)
        self.showLegend.setStatusTip('Show Legend')
        self.showLegend.triggered.connect(self.show_Legend)
        if self.view.scene().show_towns:
            self.showTowns = QtGui.QAction(QtGui.QIcon(self.check_icon), 'Show Towns', self)
        else:
            self.showTowns = QtGui.QAction(QtGui.QIcon('blank.png'), 'Show Towns', self)
        self.showTowns.setStatusTip('Show Towns')
        self.showTowns.triggered.connect(self.show_Towns)
        if self.view.scene().existing_grid:
            self.showOldg = QtGui.QAction(QtGui.QIcon(self.check_icon), 'Show Existing Grid', self)
        else:
            self.showOldg = QtGui.QAction(QtGui.QIcon('blank.png'), 'Show Existing Grid', self)
        self.showOldg.setShortcut('Ctrl+H')
        self.showOldg.setStatusTip('Show Existing Grid')
        self.showOldg.triggered.connect(self.show_OldGrid)
        if self.view.scene().existing_grid2:
            self.showOldg2 = QtGui.QAction(QtGui.QIcon(self.check_icon), 'Show Existing Grid2', self)
            self.showOldg2.setShortcut('Ctrl+2')
            self.showOldg2.setStatusTip('Show Existing Grid 2')
            self.showOldg2.triggered.connect(self.show_OldGrid2)
        self.hideTrace = QtGui.QAction(QtGui.QIcon('blank.png'), 'Clear Grid Trace', self)
        self.hideTrace.setStatusTip('Clear Grid Trace')
        self.hideTrace.triggered.connect(self.clear_Trace)
        self.showGrid = QtGui.QAction(QtGui.QIcon(self.grid_icon), 'Show Grid Line', self)
        self.showGrid.setStatusTip('Show Grid')
        self.showGrid.setShortcut('Ctrl+J')
        self.showGrid.triggered.connect(self.show_Grid)
        self.refreshGrid = QtGui.QAction(QtGui.QIcon('refresh.png'), 'Refresh Grid', self)
        self.refreshGrid.setStatusTip('Refresh Grid')
        self.refreshGrid.triggered.connect(self.refresh_Grid)
        self.goTo = QtGui.QAction(QtGui.QIcon('arrow.png'), 'Go to Station', self)
        self.goTo.setShortcut('Ctrl+G')
        self.goTo.setStatusTip('Locate specific Station')
        self.goTo.triggered.connect(self.go_To)
        self.goToTown = QtGui.QAction(QtGui.QIcon('arrowt.png'), 'Go to Town', self)
        self.goToTown.setShortcut('Ctrl+T')
        self.goToTown.triggered.connect(self.go_ToTown)
        self.saveView = QtGui.QAction(QtGui.QIcon('camera.png'), 'Save View', self)
        self.saveView.setShortcut('Ctrl+V')
        self.saveView.triggered.connect(self.save_View)
        viewMenu = menubar.addMenu('&View')
        viewMenu.addAction(self.showCapacity)
        viewMenu.addAction(self.showGeneration)
        viewMenu.addAction(self.showName)
        viewMenu.addAction(self.showFossil)
        viewMenu.addAction(self.showRuler)
        viewMenu.addAction(self.showLegend)
        viewMenu.addAction(self.showTowns)
        viewMenu.addAction(self.showOldg)
        if self.view.scene().existing_grid2:
            viewMenu.addAction(self.showOldg2)
        viewMenu.addAction(self.hideTrace)
        viewMenu.addAction(self.showGrid)
        viewMenu.addAction(self.refreshGrid)
        viewMenu.addAction(self.goTo)
        viewMenu.addAction(self.goToTown)
        if self.view.scene().load_centre is not None:
            self.goToLoad = QtGui.QAction(QtGui.QIcon('arrow.png'), 'Go to Load Centre', self)
            self.goToLoad.setShortcut('Ctrl+M')
            self.goToLoad.triggered.connect(self.go_ToLoad)
            viewMenu.addAction(self.goToLoad)
        viewMenu.addAction(self.saveView)
        self.editIni = QtGui.QAction(QtGui.QIcon('edit.png'), 'Edit Preferences File', self)
        self.editIni.setShortcut('Ctrl+E')
        self.editIni.setStatusTip('Edit Preferences')
        self.editIni.triggered.connect(self.editIniFile)
        self.editColour = QtGui.QAction(QtGui.QIcon('rainbow-icon.png'), 'Edit Colours', self)
        self.editColour.setShortcut('Ctrl+U')
        self.editColour.setStatusTip('Edit Colours')
        self.editColour.triggered.connect(self.editColours)
        self.editSect = QtGui.QAction(QtGui.QIcon('arrow.png'), 'Edit Section', self)
        self.editSect.setStatusTip('Edit Preferences Section')
        self.editSect.triggered.connect(self.editSects)
        self.editTech = QtGui.QAction(QtGui.QIcon('blank.png'), 'Edit Technologies', self)
        self.editTech.setStatusTip('Edit Technologies')
        self.editTech.triggered.connect(self.editTechs)
        self.showDtable = QtGui.QAction(QtGui.QIcon(self.grid_icon), 'Dispatchable Lines Table', self)
        self.showDtable.triggered.connect(self.showStables)
        self.showStable = QtGui.QAction(QtGui.QIcon(self.grid_icon), 'Standard Lines Table', self)
        self.showStable.triggered.connect(self.showStables)
        self.showLtable = QtGui.QAction(QtGui.QIcon('blank.png'), 'Line Costs', self)
        self.showLtable.triggered.connect(self.showStables)
        self.showSStable = QtGui.QAction(QtGui.QIcon('blank.png'), 'Substation Costs', self)
        self.showSStable.triggered.connect(self.showStables)
        editMenu = menubar.addMenu('P&references')
        editMenu.addAction(self.editIni)
        editMenu.addAction(self.editColour)
        editMenu.addAction(self.editSect)
        editMenu.addAction(self.editTech)
        editMenu.addAction(self.showDtable)
        editMenu.addAction(self.showStable)
        editMenu.addAction(self.showLtable)
        editMenu.addAction(self.showSStable)
        windowMenu = menubar.addMenu('&Windows')
        self.showResource = QtGui.QAction(QtGui.QIcon('grid.png'), 'Show Resource Grid', self)
        self.showResource.setStatusTip('Resource Grid')
        self.showResource.setShortcut('Ctrl+B')
        self.showResource.triggered.connect(self.show_Resource)
        self.showFloatMenu = QtGui.QAction(QtGui.QIcon('list.png'), 'Show Floating Menu', self)
        self.showFloatMenu.setStatusTip('Floating Menu')
        self.showFloatMenu.triggered.connect(self.show_FloatMenu)
        self.showFloatLegend = QtGui.QAction(QtGui.QIcon('list.png'), 'Show Floating Legend', self)
        self.showFloatLegend.setStatusTip('Floating Menu')
        self.showFloatLegend.triggered.connect(self.show_FloatLegend)
        self.showFloatStatus = QtGui.QAction(QtGui.QIcon('log.png'), 'Show Status Window', self)
        self.showFloatStatus.setStatusTip('Status Window')
        self.showFloatStatus.triggered.connect(self.show_FloatStatus)
        credit = QtGui.QAction(QtGui.QIcon('about.png'), 'Credits', self)
        credit.setShortcut('F2')
        credit.setStatusTip('Credits')
        credit.triggered.connect(self.showCredits)
        windowMenu.addAction(credit)
        windowMenu.addAction(self.showResource)
        if self.years is not None:
            subWindowMenu = windowMenu.addMenu('Resource for year')
            for year in self.years:
                subResource = QtGui.QAction(QtGui.QIcon('grid.png'), year, self)
                subResource.triggered.connect(self.show_Resource)
                subWindowMenu.addAction(subResource)
        windowMenu.addAction(self.showFloatLegend)
        windowMenu.addAction(self.showFloatMenu)
        windowMenu.addAction(self.showFloatStatus)
        self.credits = None
        self.resource = None
        self.floatmenu = None
        self.floatlegend = None
        self.floatstatus = None
        utilities = ['getmap', 'indexweather', 'makegrid', 'makerainfall2', 'makeweather2', 'updateswis']
        utilini = [False, True, True, False, False, True]
        utilicon = ['map.png', 'list.png', 'grid.png', 'rain.png', 'weather.png', 'list.png']
        spawns = []
        icons = []
        if sys.platform == 'win32' or sys.platform == 'cygwin':
            for i in range(len(utilities)):
                if 'rain' in utilities[i] and not self.rainfall:
                    continue
                if 'swis' in utilities[i] and 'swis' not in self.view.scene().model_name.lower():
                    continue
                if os.path.exists(utilities[i] + '.exe'):
                    if utilini[i] and len(sys.argv) > 1:
                        spawns.append([utilities[i] + '.exe', sys.argv[1]])
                    else:
                        spawns.append(utilities[i] + '.exe')
                    icons.append(utilicon[i])
                else:
                    if os.path.exists(utilities[i] + '.py'):
                        if utilini[i] and len(sys.argv) > 1:
                            spawns.append([utilities[i] + '.py', sys.argv[1]])
                        else:
                            spawns.append(utilities[i] + '.py')
                        icons.append(utilicon[i])
        else:
            for i in range(len(utilities)):
                if 'rain' in utilities[i] and not self.rainfall:
                    continue
                if 'swis' in utilities[i] and 'swis' not in self.view.scene().model_name.lower():
                    continue
                if os.path.exists(utilities[i] + '.py'):
                    if utilini[i] and len(sys.argv) > 1:
                        spawns.append([utilities[i] + '.py', sys.argv[1]])
                    else:
                        spawns.append(utilities[i] + '.py')
                    icons.append(utilicon[i])
        if len(spawns) > 0:
            spawnitem = []
            spawnMenu = menubar.addMenu('&Tools')
            for i in range(len(spawns)):
                if type(spawns[i]) is list:
                    who = spawns[i][0][:spawns[i][0].find('.')]
                else:
                    who = spawns[i][:spawns[i].find('.')]
                spawnitem.append(QtGui.QAction(QtGui.QIcon(icons[i]), who, self))
                spawnitem[-1].triggered.connect(partial(self.spawn, spawns[i]))
                spawnMenu.addAction(spawnitem[-1])
        help = QtGui.QAction(QtGui.QIcon('help.png'), 'Help', self)
        help.setShortcut('F1')
        help.setStatusTip('Help')
        help.triggered.connect(self.showHelp)
        about = QtGui.QAction(QtGui.QIcon('about.png'), 'About', self)
        about.setShortcut('Ctrl+I')
        about.setStatusTip('About')
        about.triggered.connect(self.showAbout)
        helpMenu = menubar.addMenu('&Help')
        helpMenu.addAction(help)
        helpMenu.addAction(about)

    def editIniFile(self):
        dialr = EdtDialog(self.config_file)
        dialr.exec_()
        self.get_config()   # refresh config values
        comment = self.config_file + ' edited. Reload may be required.'
        self.view.emit(SIGNAL('statusmsg'), comment)

    def changeColours(self, new_color, elements):
        for el in elements:
            if el.pen().color().name() != '#000000':
                el.setPen(QtGui.QColor(new_color))
            if el.brush().color().name() != '#000000':
                el.setBrush(QtGui.QColor(new_color))

    def editColours(self):
        dialr = Colours()
        dialr.exec_()
       # refresh some config values
        config = ConfigParser.RawConfigParser()
        config.read(self.config_file)
        technologies = config.get('Power', 'technologies')
        technologies = technologies.split(' ')
        try:
            map = config.get('Map', 'map_choice')
        except:
            map = ''
        try:
            colours0 = config.items('Colors')
        except:
            pass
        if map != '':
            try:
                colours1 = config.items('Colors' + map)
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
            if item in technologies or (item[:6] == 'fossil' and item != 'fossil_name'):
                itm = item.replace('_', ' ').title()
                itm = itm.replace('Pv', 'PV')
            else:
                itm = item
            if itm == 'ruler':
                continue
            if colour != self.view.scene().colors[itm]:
                self.changeColours(colour, self.view.scene()._stationCircles[itm])
        comment = 'Colours edited. Reload may be required.'
        self.view.emit(SIGNAL('statusmsg'), comment)

    def editTechs(self):
        EditTech(self.scenarios)
        comment = 'Technologies edited. Reload may be required.'
        self.view.emit(SIGNAL('statusmsg'), comment)

    def editSects(self):
        config = ConfigParser.RawConfigParser()
        config.read(self.config_file)
        sections = sorted(config.sections())
        menu = QtGui.QMenu()
        stns = []
        for section in sections:
            stns.append(menu.addAction(section))
            x = self.frameGeometry().left() + 50
            y = self.frameGeometry().y() + 50
        action = menu.exec_(QtCore.QPoint(x, y))
        if action != None:
            section = str(action.text())
            EditSect(section, self.scenarios)
            comment = section + ' Section edited. Reload may be required.'
            self.view.emit(SIGNAL('statusmsg'), comment)

    def showStables(self):
        fields = ['capacity', 'cost per km', 'lines']
        if 'Substation' in self.sender().text():
            field = self.view.scene().lines.substation_costs
            fields = ['line', 'cost']
        elif 'Costs' in self.sender().text():
            field = self.view.scene().lines.line_costs
            fields = ['line', 'cost per km']
        elif 'Dispatchable' in self.sender().text():
            field = self.view.scene().lines.d_line_table
        else:
            field = self.view.scene().lines.s_line_table
        dialog = displaytable.Table(field, fields=fields, title=self.sender().text(),
                 save_folder=self.scenarios)
        dialog.exec_()
        comment = self.sender().text() + ' Displayed.'
        self.view.emit(SIGNAL('statusmsg'), comment)

    def mapToLonLat(self, p):
        p = self.mapToScene(p)
        return self.scene().mapToLonLat(p)

    def mapFromLonLat(self, p):
        p = self.view.scene().mapFromLonLat(p)
        return p

    def editDescription(self, scenario):
        if scenario == False:   # called from sub menu
            scenario = self.sender().text()
        if scenario == 'Existing':
            return
        for i in range(len(self.view.scene()._scenarios)):
            if self.view.scene()._scenarios[i][0] == scenario:
                self.view.scene()._scenarios[i][2] = Description.getDescription(self.view.scene()._scenarios[i][0],
                                                     self.view.scene()._scenarios[i][2])   # get a description
                self.view.scene()._scenarios[i][1] = True
                break
        comment = 'Edited description for %s' % (scenario)
        self.view.emit(SIGNAL('statusmsg'), comment)

    def removeScenario(self, scenario):
        reshow_float = False
        if scenario == False:   # called from sub menu
            scenario = self.sender().text()
            try:
                self.subMenu.removeAction(self.sender())
            except:
                for act in self.subMenu.actions():
                    if act.text() == self.sender().text():
                        self.subMenu.removeAction(act)
                        reshow_float = True
                        break
            for act in self.subMenu2.actions():
                if act.text() == self.sender().text():
                    self.subMenu2.removeAction(act)
                    break
        for i in range(len(self.view.scene()._stations.stations) - 1, -1, -1):
            if self.view.scene()._stations.stations[i].scenario == scenario:
                self.delStation(self.view.scene()._stations.stations[i])
                del self.view.scene()._stations.stations[i]
        self.view.scene().refreshGrid()
        for i in range(len(self.view.scene()._scenarios)):
            if self.view.scene()._scenarios[i][0] == scenario:
                del self.view.scene()._scenarios[i]
                break
        if scenario == 'Existing':
            self.fileMenu.addAction(self.addExist)
            reshow_float = True
        self.view.clear_Trace()
        if reshow_float:
            self.reshow_FloatMenu()
        self.reshow_FloatLegend()
        if self.floatstatus:
            self.floatstatus.emit(SIGNAL('scenarios'), self.view.scene()._scenarios)
        comment = 'Removed scenario %s' % (scenario)
        self.view.emit(SIGNAL('statusmsg'), comment)

    def delStation(self, st):  # remove stations graphic items
        if self.view.scene()._current_name.text() == st.name:
            self.view.scene()._current_name.setText('')
        try:  # ignore error to cater for duplicate station names in different scenarios
            for itm in self.view.scene()._stationGroups[st.name]:
             #    if isinstance(self._stationGroups[self.lines.lines[i].name][j], QtGui.QGraphicsLineItem):
                try:
                    self.view.scene().removeItem(itm)  # here's the error
                except:
                    self.scene().removeItem(itm)
            del self.view.scene()._stationGroups[st.name]
        except:
            pass
        for i in range(len(self.view.scene().lines.lines) - 1, -1, -1):
            if self.view.scene().lines.lines[i].name == st.name:
                del self.view.scene().lines.lines[i]

    def saveScenariod(self):
        for i in range(len(self.view.scene()._scenarios)):
            if self.view.scene()._scenarios[i][1]:
                self.view.scene()._scenarios[i][2] = Description.getDescription(self.view.scene()._scenarios[i][0],
                                                     self.view.scene()._scenarios[i][2])  # get a description
        self.saveScenario()

    def saveScenario(self):
        comment = 'Saved scenarios: '
        for i in range(len(self.view.scene()._scenarios)):
            if self.view.scene()._scenarios[i][1]:
                for stn in self.view.scene()._stations.stations:
                    if stn.scenario == self.view.scene()._scenarios[i][0]:
                        break
                else:  # don't save empty scenario
                    continue
                self.writeStations(self.view.scene()._scenarios[i][0], self.view.scene()._scenarios[i][2])
                comment += self.view.scene()._scenarios[i][0] + ' '
                self.view.scene()._scenarios[i][1] = False
        self.view.emit(SIGNAL('statusmsg'), comment)

    def saveAsScenario(self):
        if self.new_scenario != '':
            prefix = self.new_scenario
            suffix = ''
            new = True
        else:
            i = self.scenarios_filter.find('*')
            prefix = self.scenarios_filter[:i]
            suffix = self.scenarios_filter[i + 1:]
            new = False
        fname = QtGui.QFileDialog.getSaveFileName(self, 'Save scenario file',
                self.scenarios + prefix, 'Scenarios (' + self.scenarios_filter + ')')
        if fname != '':
            fname = str(fname)
            i = fname.rfind('/')
            save_as = fname[i + 1:]
            if not new:
                if save_as[:len(prefix)] != prefix:
                    save_as = prefix + save_as
                if save_as[-len(suffix):] != suffix:
                    save_as = save_as + suffix
            desc = ''
            for i in range(len(self.view.scene()._scenarios)):
                if self.view.scene()._scenarios[i][0] == save_as:
                    desc = self.view.scene()._scenarios[i][2]
                    break
            description = Description.getDescription(save_as, desc)  # get a description
            for i in range(len(self.view.scene()._stations.stations)):
                if self.view.scene()._stations.stations[i].scenario != 'Existing':
                    self.view.scene()._stations.stations[i].scenario = save_as
            self.writeStations(save_as, description)
            exist = False
            for i in range(len(self.view.scene()._scenarios) - 1, -1, -1):
                if self.view.scene()._scenarios[i][0] == 'Existing':
                    exist = True
                else:
                    del self.view.scene()._scenarios[i]
            self.view.scene()._scenarios.append([save_as, False, description])
            self.subMenu.clear()
            subFile = QtGui.QAction(QtGui.QIcon('minus.png'), save_as, self)
            subFile.setStatusTip('Remove Scenario ' + save_as)
            subFile.triggered.connect(self.removeScenario)
            self.subMenu.addAction(subFile)
            if exist:
                subFile = QtGui.QAction(QtGui.QIcon('minus.png'), 'Existing', self)
                subFile.setStatusTip('Remove Scenario Existing')
                subFile.triggered.connect(self.removeScenario)
                self.subMenu.addAction(subFile)
            self.subMenu2.clear()
            subFile2 = QtGui.QAction(QtGui.QIcon('edit.png'), save_as, self)
            subFile2.setStatusTip('Edit Description for ' + save_as)
            subFile2.triggered.connect(self.editDescription)
            self.subMenu2.addAction(subFile2)
            if self.floatstatus:
                self.floatstatus.emit(SIGNAL('scenarios'), self.view.scene()._scenarios)
            self.reshow_FloatMenu()
            comment = 'Saved as scenario: ' + save_as
            self.view.emit(SIGNAL('statusmsg'), comment)

    def addExisting(self):
        stations = Stations()
        for st in stations.stations:
            self.view.scene()._stations.stations.append(st)
            self.view.scene().addStation(st)
        self.view.scene()._scenarios.append(['Existing', False, 'Existing stations'])
        try:
            self.fileMenu.removeAction(self.sender())
        except:
            for act in self.fileMenu.actions():
                if act.text() == self.sender().text():
                    self.fileMenu.removeAction(act)
                    break
        subFile = QtGui.QAction(QtGui.QIcon('minus.png'), 'Existing', self)
        subFile.setStatusTip('Remove Scenario Existing')
        subFile.triggered.connect(self.removeScenario)
        self.subMenu.addAction(subFile)
        self.reshow_FloatMenu()
        self.reshow_FloatLegend()
        if self.floatstatus:
            self.floatstatus.emit(SIGNAL('scenarios'), self.view.scene()._scenarios)
        self.view.emit(SIGNAL('statusmsg'), 'Added Existing stations')

    def addScenario(self):
        fname = QtGui.QFileDialog.getOpenFileName(self, 'Open scenario file',
                self.scenarios, 'Scenarios (' + self.scenarios_filter + ')')
        fname = str(fname)
        if os.path.exists(fname):
            if self.sender().text() == 'Open':
                for i in range(len(self.view.scene()._scenarios) - 1, -1, -1):
                    scen = self.view.scene()._scenarios[i][0]
                    self.removeScenario(scen)
                self.subMenu.clear()
                self.subMenu2.clear()
            self.view.scene()._setupScenario(fname)
            i = fname.rfind('/')
            scen_filter = fname[i + 1:]
            comment = 'Added scenario %s' % (scen_filter)
            subFile = QtGui.QAction(QtGui.QIcon('minus.png'), scen_filter, self)
            subFile.setStatusTip('Remove Scenario ' + scen_filter)
            subFile.triggered.connect(self.removeScenario)
            self.subMenu.addAction(subFile)
            subFile2 = QtGui.QAction(QtGui.QIcon('edit.png'), scen_filter, self)
            subFile2.setStatusTip('Edit Description for ' + scen_filter)
            subFile2.triggered.connect(self.editDescription)
            self.subMenu2.addAction(subFile2)
            self.reshow_FloatMenu()
            self.reshow_FloatLegend()
            if self.floatstatus:
                self.floatstatus.emit(SIGNAL('scenarios'), self.view.scene()._scenarios)
            self.altered_stations = True
            self.view.emit(SIGNAL('statusmsg'), comment)

    def center(self):
        frameGm = self.frameGeometry()
        screen = QtGui.QApplication.desktop().screenNumber(QtGui.QApplication.desktop().cursor().pos())
        centerPoint = QtGui.QApplication.desktop().availableGeometry(screen).center()
        frameGm.moveCenter(centerPoint)
        self.move(frameGm.topLeft())

    def showHelp(self):
        if self.leave_help_open:
            webbrowser.open_new(self.helpfile)
        else:
            dialog = displayobject.AnObject(QtGui.QDialog(), self.helpfile, title='Help')
            dialog.exec_()

    def showCredits(self, initial=False):
        if self.credits is None:
            self.credits = Credits(initial)
            self.credits.setWindowModality(QtCore.Qt.WindowModal)
            self.credits.setWindowFlags(self.credits.windowFlags() |
                         QtCore.Qt.WindowSystemMenuHint |
                         QtCore.Qt.WindowMinMaxButtonsHint)
            self.credits.procStart.connect(self.getCredits)
            self.credits.show()
     #        self.credits.exec_()

    @QtCore.pyqtSlot(str)
    def getCredits(self, text):
        if text == 'goodbye':
            self.credits = None
        elif text == 'help':
            self.showHelp()
        elif text == 'about':
            self.showAbout()

    @QtCore.pyqtSlot(str)
    def startCredit(self):
        self.showCredits()

    def show_FloatStatus(self):
        if self.floatstatus is None:
            self.floatstatus = FloatStatus(self, self.view.scene().scenarios, self.view.scene()._scenarios)
            self.floatstatus.setWindowModality(QtCore.Qt.WindowModal)
            self.floatstatus.setWindowFlags(self.floatstatus.windowFlags() |
                         QtCore.Qt.WindowSystemMenuHint |
                         QtCore.Qt.WindowMinMaxButtonsHint)
            self.floatstatus.procStart.connect(self.getStatus)
            if self.log_status:
                self.connect(self.floatstatus, SIGNAL('log'), self.floatstatus.log)
            self.connect(self.floatstatus, SIGNAL('scenarios'), self.floatstatus.updateScenarios)
            self.floatstatus.show()
            self.activateWindow()
     #        self.credits.exec_()

    @QtCore.pyqtSlot(str)
    def getStatus(self, text):
        if text == 'goodbye':
            self.floatstatus = None

    def showAbout(self):
        dialog = displayobject.AnObject(QtGui.QDialog(), self.aboutfile, title='About SENs SAM Model')
        dialog.exec_()

    def setStatusText(self, text):
        if text == self.statusBar().currentMessage():
            return
        if self.floatstatus and self.log_status:
            if text[0] != '(':
                self.floatstatus.emit(SIGNAL('log'), '%s. %s' %
                            (str(QtCore.QDateTime.toString(QtCore.QDateTime.currentDateTime(),
                            'hh:mm:ss')), text))
      #   self.statusBar().clearMessage()
        self.statusBar().showMessage(text)

    def go_To(self):
        menu = QtGui.QMenu()
        stns = []
        stns_to_show = 0
        for station in self.view.scene()._stations.stations:
             if station.technology[:6] == 'Fossil' and not self.view.scene().show_fossil:
                 continue
             stns_to_show += 1
        try:
            self.icons
        except:
            self.icons = Icons()
        if stns_to_show > 25:
            submenus = []
            submenus.append(QtGui.QMenu('A...'))
            titl = ''
            ctr = 0
            last_stn = ''
            for station in sorted(self.view.scene()._stations.stations, key=lambda station: station.name):
                if station.technology[:6] == 'Fossil' and not self.view.scene().show_fossil:
                    continue
                if titl == '':
                    titl = station.name + ' to '
                icon = self.icons.getIcon(station.technology)
                stns.append(submenus[-1].addAction(QtGui.QIcon(icon), station.name))
                stns[-1].setIconVisibleInMenu(True)
                ctr += 1
                if ctr > 25:
                    titl += station.name
                    submenus[-1].setTitle(titl)
                    titl = ''
                    menu.addMenu(submenus[-1])
                    submenus.append(QtGui.QMenu(station.name[0] + '...'))
                    ctr = 0
                last_stn = station.name
            titl += last_stn
            submenus[-1].setTitle(titl)
            menu.addMenu(submenus[-1])
        else:
            for station in sorted(self.view.scene()._stations.stations, key=lambda station: station.name):
                if station.technology[:6] == 'Fossil':
                    if not self.view.scene().show_fossil:
                        continue
                icon = self.icons.getIcon(station.technology)
                stns.append(menu.addAction(QtGui.QIcon(icon), station.name))
                stns[-1].setIconVisibleInMenu(True)
        x = self.frameGeometry().left() + 50
        y = self.frameGeometry().y() + 50
        action = menu.exec_(QtCore.QPoint(x, y))
        if action != None:
            station = self.view.scene()._stations.Get_Station(action.text())
            go_to = self.mapFromLonLat(QtCore.QPointF(station.lon, station.lat))
            self.view.centerOn(go_to)
            if not self.view.scene().show_station_name:
                self.view.scene()._current_name.setText(station.name)
                pp = self.mapFromLonLat(QtCore.QPointF(station.lon, station.lat))
                self.view.scene()._current_name.setPos(pp + QtCore.QPointF(1.5, -0.5))
                if station.technology[:6] == 'Fossil':
                    self.view.scene()._current_name.setBrush(QtGui.QColor(self.view.scene().colors['fossil_name']))
                else:
                    self.view.scene()._current_name.setBrush(QtGui.QColor(self.view.scene().colors['station_name']))
# highlight grid line
            comment = '(%s,%s) Centred on station %s' % ('{:0.4f}'.format(station.lat), '{:0.4f}'.format(station.lon),
                          station.name)
            self.view.emit(SIGNAL('statusmsg'), comment)

    def go_ToTown(self):
     #   to cater for windows I've created submenus
        menu = QtGui.QMenu()
        twns = []
        submenus = []
        submenus.append(QtGui.QMenu('A...'))
        titl = ''
        ctr = 0
        for town in sorted(self.view.scene()._towns.towns, key=lambda town: town.name):
            if titl == '':
                titl = town.name + ' to '
            twns.append(submenus[-1].addAction(town.name))
            twns[-1].setIconVisibleInMenu(True)
            ctr += 1
            if ctr > 25:
                titl += town.name
                submenus[-1].setTitle(titl)
                titl = ''
                menu.addMenu(submenus[-1])
                submenus.append(QtGui.QMenu(town.name[0] + '...'))
                ctr = 0
        titl += town.name
        submenus[-1].setTitle(titl)
        menu.addMenu(submenus[-1])
        x = self.frameGeometry().left() + 50
        y = self.frameGeometry().y() + 50
        action = menu.exec_(QtCore.QPoint(x, y))
        if action != None:
            town = self.view.scene()._towns.Get_Town(action.text())
            go_to = self.mapFromLonLat(QtCore.QPointF(town.lon, town.lat))
            self.view.centerOn(go_to)
            comment = '(%s,%s) Centred on town %s' % ('{:0.4f}'.format(town.lon), '{:0.4f}'.format(town.lat),
                          town.name)
            self.view.emit(SIGNAL('statusmsg'), comment)

    def go_ToLoad(self):
        if len(self.view.scene().load_centre) == 1:
            j = 0
        else:
            menu = QtGui.QMenu()
            ctrs = []
            for ctr in self.view.scene().load_centre:
                ctrs.append(menu.addAction(ctr[0]))
           #      ctrs[-1].setIconVisibleInMenu(True)
            x = self.frameGeometry().left() + 50
            y = self.frameGeometry().y() + 50
            action = menu.exec_(QtCore.QPoint(x, y))
            if action != None:
                for j in range(len(self.view.scene().load_centre)):
                    if action.text() == self.view.scene().load_centre[j][0]:
                        break
        if j < len(self.view.scene().load_centre):
            go_to = self.mapFromLonLat(QtCore.QPointF(self.view.scene().load_centre[j][2],
                    self.view.scene().load_centre[j][1]))
            self.view.centerOn(go_to)
            comment = '(%s,%s) Centred on %s Load Centre' % (
                      '{:0.4f}'.format(self.view.scene().load_centre[j][1]),
                      '{:0.4f}'.format(self.view.scene().load_centre[j][2]),
                      self.view.scene().load_centre[j][0])
            self.view.emit(SIGNAL('statusmsg'), comment)

    def save_View(self):
        outputimg = QtGui.QPixmap(self.view.width(), self.view.height())
        painter = QtGui.QPainter(outputimg)
        targetrect = QtCore.QRectF(0, 0, self.view.width(), self.view.height())
        sourcerect = QtCore.QRect(0, 0, self.view.width(), self.view.height())
        self.view.render(painter, targetrect, sourcerect)
        fname = self.new_scenario[:self.new_scenario.rfind('.')] + '.png'
        fname = QtGui.QFileDialog.getSaveFileName(self, 'Save image file',
                self.scenarios + fname, 'Image Files (*.png *.jpg *.bmp)')
        if fname != '':
            fname = str(fname)
            i = fname.rfind('.')
            if i < 0:
                fname = fname + '.png'
                i = fname.rfind('.')
            outputimg.save(fname, fname[i + 1:])
            try:
                comment = 'View saved to ' + fname[fname.rfind('/') + 1:]
            except:
                comment = 'View saved to ' + fname
            self.view.emit(SIGNAL('statusmsg'), comment)
        painter.end()

    def spawn(self, who):
        if type(who) is list:
            if os.path.exists(who[0]):
                if sys.platform == 'win32' or sys.platform == 'cygwin':
                    if who[0][-3:] == '.py':
                        pid = subprocess.Popen([who], shell=True).pid
                    else:
                        pid = subprocess.Popen([who[0], who[1]]).pid
                else:
                    pid = subprocess.Popen(['python', who[0], who[1]]).pid
                self.view.emit(SIGNAL('statusmsg'), who[0] + ' invoked')
        else:
            if os.path.exists(who):
                if sys.platform == 'win32' or sys.platform == 'cygwin':
                    if who[-3:] == '.py':
                        pid = subprocess.Popen([who], shell=True).pid
                    else:
                        pid = subprocess.Popen([who]).pid
                else:
                    pid = subprocess.Popen(['python', who]).pid
            self.view.emit(SIGNAL('statusmsg'), who + ' invoked')
        return

    def popup(self, pos):
        def check_scenario():
            self.altered_stations = True
            station = self.view.scene()._stations.stations[-1]  # point to last station
            for i in range(len(self.view.scene()._scenarios)):
                if self.view.scene()._scenarios[i][0] == station.scenario:
                    self.view.scene()._scenarios[i][1] = True
                    break
            else:
                self.view.scene()._scenarios.append([station.scenario, False, ''])
                subFile = QtGui.QAction(QtGui.QIcon('minus.png'), station.scenario, self)
                subFile.setStatusTip('Remove Scenario ' + station.scenario)
                subFile.triggered.connect(self.removeScenario)
                self.subMenu.addAction(subFile)
                subFile2 = QtGui.QAction(QtGui.QIcon('edit.png'), station.scenario, self)
                subFile2.setStatusTip('Edit Description for ' + station.scenario)
                subFile2.triggered.connect(self.editDescription)
                self.subMenu2.addAction(subFile2)
                self.reshow_FloatMenu()
                if self.floatstatus:
                    self.floatstatus.emit(SIGNAL('scenarios'), self.view.scene()._scenarios)
                self.altered_stations = True
                self.view.scene()._scenarios[-1][1] = True
            self.reshow_FloatLegend()

        where = self.view.mapToLonLat(pos)
        menu = QtGui.QMenu()
        station = None
        if len(self.view.scene()._stations.stations) > 0:  # some stations
            station, st_dist = self.view.scene()._stations.Nearest(where.y(), where.x(),
            distance=True, fossil=self.view.scene().show_fossil)
            titl = 'Nearest station: %s (%s MW; %s Km away)' % (station.name, '{:0.0f}'.format(station.capacity),
                   '{:0.0f}'.format(st_dist))
         #    act1 = 'Run SAM API (%s)' % model
            if station.technology == 'Hydro' or station.technology == 'Wave' or station.technology[:5] == 'Other':
                act2 = 'Run Power Model for %s' % station.name
            else:
                act2 = 'Run SAM Power Model for %s' % station.name
            act3 = 'Center view on %s' % station.name
            if station.scenario != 'Existing' and station.technology[:6] != 'Fossil':
                act4 = 'Show/Edit details for %s' % station.name
            else:
                act4 = 'Show details for %s' % station.name
            act5 = 'Move %s' % station.name
            act6 = 'Delete %s' % station.name
            act8 = 'Copy %s' % station.name
            act12 = 'Edit grid line for %s' % station.name
            act13 = 'Trace grid for %s' % station.name
        try:
            town, town_dist = self.view.scene()._towns.Nearest(where.y(), where.x(), distance=True)
            ttitl = 'Nearest town: %s (%s Km away)' % (town.name, '{:0.0f}'.format(town_dist))
            act9 = 'Show details for %s' % town.name
        except:
            ttitl = 'No towns found'
        act7 = 'Add Station at %s %s' % ('{:0.4f}'.format(where.y()), '{:0.4f}'.format(where.x()))
        act10 = 'Position ruler here'
        act11 = 'Position legend here'
        act14 = 'Show weather for here'
        if station is not None:  # a stations
            noAction = menu.addAction(titl)
            if station.technology[:6] != 'Fossil':
                run1Action = menu.addAction(QtGui.QIcon('power.png'), act2)
                run1Action.setIconVisibleInMenu(True)
                cpyAction = menu.addAction(QtGui.QIcon('copy.png'), act8)
                cpyAction.setIconVisibleInMenu(True)
            ctrAction = menu.addAction(QtGui.QIcon('zoom.png'), act3)
            ctrAction.setIconVisibleInMenu(True)
            stnAction = menu.addAction(QtGui.QIcon('info_green.png'), act4)
            stnAction.setIconVisibleInMenu(True)
            if station.scenario != 'Existing':
                mveAction = menu.addAction(QtGui.QIcon('move.png'), act5)
                mveAction.setIconVisibleInMenu(True)
                delAction = menu.addAction(QtGui.QIcon('minus.png'), act6)
                delAction.setIconVisibleInMenu(True)
                grdAction = menu.addAction(QtGui.QIcon('network_b.png'), act12)
                grdAction.setIconVisibleInMenu(True)
            if self.view.scene().trace_grid:
                trcAction = menu.addAction(QtGui.QIcon('line.png'), act13)
                trcAction.setIconVisibleInMenu(True)
            else:
                trcAction = ''
        addsAction = menu.addAction(QtGui.QIcon('plus.png'), act7)
        addsAction.setIconVisibleInMenu(True)
        sunAction = menu.addAction(QtGui.QIcon('weather.png'), act14)
        sunAction.setIconVisibleInMenu(True)
        notAction = menu.addAction(ttitl)
        if ttitl[:2] != 'No':
            twnAction = menu.addAction(QtGui.QIcon('info.png'), act9)
            twnAction.setIconVisibleInMenu(True)
        rulAction = menu.addAction(QtGui.QIcon('ruler.png'), act10)
        rulAction.setIconVisibleInMenu(True)
        legAction = menu.addAction(QtGui.QIcon('list.png'), act11)
        legAction.setIconVisibleInMenu(True)
        action = menu.exec_(self.mapToGlobal(pos))
        if action == None:
            return
        try:
            if action == noAction or action == notAction:
                return
        except:
            pass
        if action == addsAction:
            try:
                new_station = Station(town.name + ' Station', '', where.y(), where.x(),
                              0.0, '', 0.0, 0, 0.0, self.new_scenario)
            except:
                new_station = Station('Station ' + str(where.y()) + ' ' + str(where.x()), '', where.y(), where.x(),
                              0.0, '', 0.0, 0, 0.0, self.new_scenario)
            dialog = newstation.AnObject(QtGui.QDialog(), new_station, scenarios=self.view.scene()._scenarios)
            dialog.exec_()
            new_station = dialog.getValues()
            if new_station is not None:
                name_ok = False
                new_name = new_station.name
                ctr = 0
                while not name_ok:
                    for i in range(len(self.view.scene()._stations.stations)):
                        if self.view.scene()._stations.stations[i].name == new_name:
                            ctr += 1
                            new_name = new_station.name + ' ' + str(ctr)
                            break
                    else:
                        name_ok = True
                if new_name != new_station.name:
                    new_station.name = new_name
                self.view.scene()._stations.stations.append(new_station)
                self.view.scene().addStation(new_station)
                where.y = new_station.lat
                where.x = new_station.lon
                comment = 'Added station %s (%s MW at %s,%s) ' % (new_station.name,
                          '{:0.0f}'.format(new_station.capacity), '{:0.4f}'.format(new_station.lat),
                          '{:0.4f}'.format(new_station.lon))
                check_scenario()
                self.view.emit(SIGNAL('statusmsg'), comment)
        elif action == rulAction:
            if self.view.scene().show_ruler:
                self.view.hide_Ruler()
            self.view.show_Ruler(self.view.scene().ruler, self.view.scene().ruler_ticks, pos)
            self.showRuler.setIcon(QtGui.QIcon(self.check_icon))
            self.view.scene().show_ruler = True
            self.view.emit(SIGNAL('statusmsg'), 'Scale Ruler Toggled On')
        elif action == legAction:
            if self.view.scene().show_legend:
                self.view.hide_Legend()
            self.view.show_Legend(pos, where=where)
            self.showLegend.setIcon(QtGui.QIcon(self.check_icon))
            self.view.scene().show_legend = True
            self.view.emit(SIGNAL('statusmsg'), 'Legend Toggled On')
        elif action == sunAction:
            PlotWeather(where.y(), where.x(), self.base_year)
            self.view.emit(SIGNAL('statusmsg'), 'Weather displayed')
        elif action == ctrAction:
            go_to = self.mapFromLonLat(QtCore.QPointF(station.lon, station.lat))
            self.view.centerOn(go_to)
        elif ttitl[:2] != 'No' and action == twnAction:
            dialog = displayobject.AnObject(QtGui.QDialog(), town)
            dialog.exec_()
        elif action == stnAction:
            if station.scenario != 'Existing':
                s_was = station.scenario
                dialog = newstation.AnObject(QtGui.QDialog(), station, scenarios=self.view.scene()._scenarios)
                p_was = self.mapFromLonLat(QtCore.QPointF(station.lon, station.lat))
                dialog.exec_()
                chg_station = dialog.getValues()
                if chg_station is not None:
                    if chg_station.name != station.name:
                        name_ok = False
                        new_name = chg_station.name
                        ctr = 0
                        while not name_ok:
                            for i in range(len(self.view.scene()._stations.stations)):
                                if self.view.scene()._stations.stations[i].name == new_name:
                                    ctr += 1
                                    new_name = chg_station.name + ' ' + str(ctr)
                                    break
                            else:
                                name_ok = True
                        if new_name != chg_station.name:
                            chg_station.name = new_name
                    self.delStation(station)
                    for atr in dir(chg_station):
                        if atr[:2] != '__' and atr[-2:] != '__':
                            setattr(station, atr, getattr(chg_station, atr))
                    self.view.scene().addStation(station)
                    self.altered_stations = True
                    for i in range(len(self.view.scene()._scenarios)):
                        if self.view.scene()._scenarios[i][0] == station.scenario \
                        or self.view.scene()._scenarios[i][0] == s_was:
                            self.view.scene()._scenarios[i][1] = True
                    comment = 'Altered station %s (%s,%s) ' % (station.name,
                              '{:0.4f}'.format(station.lat), '{:0.4f}'.format(station.lon))
                    self.view.emit(SIGNAL('statusmsg'), comment)
                    self.reshow_FloatLegend()
            else:
                dialog = displayobject.AnObject(QtGui.QDialog(), station)
                dialog.exec_()
        elif action == trcAction:
            grid_path_len = self.view.traceGrid(station)
            try:
                self.view.emit(SIGNAL('statusmsg'), 'Grid traced for %s (%s Km)' % (station.name,
                              '{:0.1f}'.format(grid_path_len)))
            except:
                self.view.emit(SIGNAL('statusmsg'), 'No Grid line for %s' % station.name)
        elif action == run1Action:
            power = PowerModel([station])
            generated = power.getValues()
            station.generation = generated[0].generation
            comment = 'Power plot completed for %s. %s MWh; CF %s' % (station.name,
                      '{:0,.1f}'.format(generated[0].generation),
                      '{:0.2f}'.format(generated[0].cf))
            self.view.emit(SIGNAL('statusmsg'), comment)
        elif action == cpyAction:
            self.view.clear_Trace()
            new_station = Station(station.name + ' 2', station.technology, station.lat,
                          station.lon, station.capacity, station.turbine, station.rotor,
                          station.no_turbines, station.area, self.new_scenario)
            dialog = newstation.AnObject(QtGui.QDialog(), new_station, scenarios=self.view.scene()._scenarios)
            dialog.exec_()
            new_station = dialog.getValues()
            if new_station is not None:
                name_ok = False
                new_name = new_station.name
                ctr = 0
                while not name_ok:
                    for i in range(len(self.view.scene()._stations.stations)):
                        if self.view.scene()._stations.stations[i].name == new_name:
                            ctr += 1
                            new_name = new_station.name + ' ' + str(ctr)
                            break
                    else:
                        name_ok = True
                    if new_name != new_station.name:
                        new_station.name = new_name
                self.view.scene()._stations.stations.append(new_station)
                self.view.scene().addStation(new_station)
                where.x = new_station.lon
                where.y = new_station.lat
                comment = 'Added station %s (%s MW at %s,%s) ' % (new_station.name,
                          '{:0.0f}'.format(new_station.capacity), '{:0.4f}'.format(new_station.lat),
                          '{:0.4f}'.format(new_station.lon))
                check_scenario()
                self.view._move_station = True
                station = self.view.scene()._stations.stations[-1]
                self.view._station_to_move = station
                self.altered_stations = True
                for i in range(len(self.view.scene()._scenarios)):
                    if self.view.scene()._scenarios[i][0] == station.scenario:
                        self.view.scene()._scenarios[i][1] = True
                        break
                self.view.emit(SIGNAL('statusmsg'), comment)
        elif action == mveAction:
            self.view.clear_Trace()
            self.view._move_station = True
            self.view._station_to_move = station
            self.altered_stations = True
            for i in range(len(self.view.scene()._scenarios)):
                if self.view.scene()._scenarios[i][0] == station.scenario:
                    self.view.scene()._scenarios[i][1] = True
                    break
            self.reshow_FloatLegend()
            comment = 'Moved station %s (%s MW at %s,%s) ' % (station.name,
                      '{:0.0f}'.format(station.capacity), '{:0.4f}'.format(station.lat),
                      '{:0.4f}'.format(station.lon))
            self.view.emit(SIGNAL('statusmsg'), comment)
        elif action == delAction:
            self.view.clear_Trace()
            p = self.mapFromLonLat(QtCore.QPointF(station.lon, station.lat))
            comment = 'Removed station %s (%s,%s) ' % (station.name,
                      '{:0.4f}'.format(station.lat), '{:0.4f}'.format(station.lon))
            for i in range(len(self.view.scene()._stations.stations) - 1, -1, -1):
                if self.view.scene()._stations.stations[i].name == station.name:
                    self.delStation(self.view.scene()._stations.stations[i])
                    for j in range(len(self.view.scene()._scenarios)):
                        if self.view.scene()._scenarios[j][0] == station.scenario:
                            self.view.scene()._scenarios[j][1] = True
                            break
                    del self.view.scene()._stations.stations[i]
                    self.view.scene().refreshGrid()
                    break
            self.altered_stations = True
            self.reshow_FloatLegend()
            self.view.emit(SIGNAL('statusmsg'), comment)
        elif action == grdAction:
            self.view.clear_Trace()
            p = self.mapFromLonLat(QtCore.QPointF(station.lon, station.lat))
            self.view._move_grid = True
            self.view._grid_start = p
            self.view._station_to_move = station
            self.altered_stations = True
            for i in range(len(self.view.scene()._scenarios)):
                if self.view.scene()._scenarios[i][0] == station.scenario:
                    self.view.scene()._scenarios[i][1] = True
                    break
            self.view.emit(SIGNAL('statusmsg'), 'Edit Grid Line for ' + station.name + '. Double-click to finish')

    def get_Power(self):
        if self.sender().text()[:4] == 'Powe':
            power = PowerModel(self.view.scene()._stations.stations)
        else:
            power = PowerModel(self.view.scene()._stations.stations, year=self.sender().text())
        generated = power.getValues()
        for stn in generated:
            station = self.view.scene()._stations.Get_Station(stn.name)
            station.generation = stn.generation
            self.view.scene().addGeneration(station)
        comment = 'Power plot completed'
        pct = power.getPct()
        if pct is not None:
            comment += ' (generation meets ' + pct[2:]
        self.view.emit(SIGNAL('statusmsg'), comment)

    def list_Stations(self):
        ctr = [0, 0]
        for st in self.view.scene()._stations.stations:
            if st.technology[:6] != 'Fossil':
                ctr[0] += 1
                if st.generation > 0:
                    ctr[1] += 1
        fields = ['name', 'technology', 'capacity']
        units = 'capacity=MW'
        sumfields = ['capacity']
        if ctr[1] > 1:
            fields.append('generation')
            units += ' generation=MWh'
            sumfields.append('generation')
        fields.append('scenario')
        dialog = displaytable.Table(self.view.scene()._stations.stations,
                 fossil=self.view.scene().show_fossil, fields=fields,
                 units=units, sumby='technology', sumfields=sumfields,
                 save_folder=self.scenarios)
        dialog.exec_()
        comment = 'Stations Displayed'
        self.view.emit(SIGNAL('statusmsg'), comment)

    def list_Grid(self):
        if self.view.scene().cost_existing:
            for i in range(self.view.scene().grid_lines):
                cost, self.view.scene().lines.lines[i].line_table = \
                        self.view.scene().lines.Line_Cost(self.view.scene().lines.lines[i].peak_load,
                        self.view.scene().lines.lines[i].peak_dispatchable)
                self.view.scene().lines.lines[i].line_cost = cost * self.view.scene().lines.lines[i].length
        for i in range(len(self.view.scene().lines.lines)):
            for j in range(len(self.view.scene().load_centre)):
                if self.view.scene().lines.lines[i].coordinates[-1] == [self.view.scene().load_centre[j][1],
                   self.view.scene().load_centre[j][2]]:
                    if self.view.scene().lines.lines[i].peak_load is not None:
                        a, self.view.scene().lines.lines[i].substation_cost, b = \
                           self.view.scene().lines.decode2(str(self.view.scene().lines.lines[i].peak_load) + '=' +
                           self.view.scene().lines.lines[i].line_table, substation=True)
                    j = -1
                    break
            if j < 0:
                cost = 0.
                if self.view.scene().lines.lines[i].peak_load is not None:
                    for key, value in self.view.scene().lines.substation_costs.iteritems():
                        if key in self.view.scene().lines.lines[i].line_table:
                            if value > cost:
                                cost = value
                if self.view.scene().lines.lines[i].connector >= 0:
                    j = self.view.scene().lines.lines[i].connector
                    if self.view.scene().lines.lines[j].peak_load is not None:
                        for key, value in self.view.scene().lines.substation_costs.iteritems():
                            if key in self.view.scene().lines.lines[j].line_table:
                                if value > cost:
                                    cost = value
                self.view.scene().lines.lines[i].substation_cost = cost
        dialog = displaytable.Table(self.view.scene().lines.lines,
                 fields=['name', 'line_table', 'length', 'line_cost', 'substation_cost', 'peak_load', 'peak_dispatchable',
                         'peak_loss', 'coordinates', 'connector'],
                 units='length=Km line_cost=$ substation_cost=$ peak_load=MW peak_dispatchable=MW peak_loss=MW',
                 sumfields=['length', 'line_cost', 'substation_cost'],
                 save_folder=self.scenarios)  # '#', 'connector',
        dialog.exec_()
        comment = 'Grid Displayed'
        self.view.emit(SIGNAL('statusmsg'), comment)

    def show_Capacity(self):
        comment = 'Capacity Circles Toggled'
        if self.view.scene().show_capacity:
            self.showCapacity.setIcon(QtGui.QIcon('blank.png'))
            self.view.scene().show_capacity = False
            self.view.scene()._capacityGroup.setVisible(False)
            self.view.scene()._fcapacityGroup.setVisible(False)
            comment += ' Off'
        else:
            self.showCapacity.setIcon(QtGui.QIcon(self.check_icon))
            self.view.scene().show_capacity = True
            self.view.scene()._capacityGroup.setVisible(True)
            if self.view.scene().show_fossil:
                self.view.scene()._fcapacityGroup.setVisible(True)
            comment += ' On'
        self.reshow_FloatLegend()
        self.view.emit(SIGNAL('statusmsg'), comment)

    def show_Generation(self):
        comment = 'Generation Circles Toggled'
        if self.view.scene().show_generation:
            self.showGeneration.setIcon(QtGui.QIcon('blank.png'))
            self.view.scene().show_generation = False
            self.view.scene()._generationGroup.setVisible(False)
            comment += ' Off'
        else:
            self.showGeneration.setIcon(QtGui.QIcon(self.check_icon))
            self.view.scene().show_generation = True
            self.view.scene()._generationGroup.setVisible(True)
            comment += ' On'
        self.reshow_FloatLegend()
        self.view.emit(SIGNAL('statusmsg'), comment)

    def show_Fossil(self):
        comment = 'Fossil-fueled Stations Toggled'
        if self.view.scene().show_fossil:
            self.showFossil.setIcon(QtGui.QIcon('blank.png'))
            self.view.scene().show_fossil = False
            self.view.scene()._fossilGroup.setVisible(False)
            self.view.scene()._fcapacityGroup.setVisible(False)
            self.view.scene()._fnameGroup.setVisible(False)
            comment += ' Off'
        else:
            self.showFossil.setIcon(QtGui.QIcon(self.check_icon))
            self.view.scene().show_fossil = True
            self.view.scene()._fossilGroup.setVisible(True)
            if self.view.scene().show_capacity:
                self.view.scene()._fcapacityGroup.setVisible(True)
            if self.view.scene().show_station_name:
                self.view.scene()._fnameGroup.setVisible(True)
            comment += ' On'
        self.reshow_FloatLegend()
        self.view.emit(SIGNAL('statusmsg'), comment)

    def show_Towns(self):
        comment = 'Towns Toggled'
        if self.view.scene().show_towns:
            self.showTowns.setIcon(QtGui.QIcon('blank.png'))
            self.view.scene().show_towns = False
            self.view.scene()._townGroup.setVisible(False)
            comment += ' Off'
        else:
            self.showTowns.setIcon(QtGui.QIcon(self.check_icon))
            self.view.scene().show_towns = True
            self.view.scene()._townGroup.setVisible(True)
            comment += ' On'
        self.view.emit(SIGNAL('statusmsg'), comment)

    def show_OldGrid(self):
        comment = 'Existing Grid Toggled'
        if self.view.scene().existing_grid:
            self.showOldg.setIcon(QtGui.QIcon('blank.png'))
            self.view.scene().existing_grid = False
            self.view.scene()._gridGroup.setVisible(False)
            comment += ' Off'
        else:
            self.showOldg.setIcon(QtGui.QIcon(self.check_icon))
            self.view.scene().existing_grid = True
            self.view.scene()._gridGroup.setVisible(True)
            comment += ' On'
        self.view.emit(SIGNAL('statusmsg'), comment)

    def show_OldGrid2(self):
        comment = 'Existing Grid2 Toggled'
        if self.view.scene().existing_grid2:
            self.showOldg2.setIcon(QtGui.QIcon('blank.png'))
            self.view.scene().existing_grid2 = False
            self.view.scene()._gridGroup2.setVisible(False)
            comment += ' Off'
        else:
            self.showOldg2.setIcon(QtGui.QIcon(self.check_icon))
            self.view.scene().existing_grid2 = True
            self.view.scene()._gridGroup2.setVisible(True)
            comment += ' On'
        self.view.emit(SIGNAL('statusmsg'), comment)

    def clear_Trace(self):
        self.view.clear_Trace()
        self.view.emit(SIGNAL('statusmsg'), 'Grid Trace cleared')

    def refresh_Grid(self):
        self.view.clear_Trace()
        self.view.scene().refreshGrid()
        self.view.emit(SIGNAL('statusmsg'), 'Grid Refreshed')

    def show_Grid(self):
        self.view.clear_Trace()
        menu = QtGui.QMenu()
        lins = []
       #  for li in range(len(self.view.scene().lines.lines)): #.grid_lines):
        for li in range(self.view.scene().grid_lines):
            lins.append(menu.addAction(self.view.scene().lines.lines[li].name))
            lins[-1].setIconVisibleInMenu(True)
        for li in range(self.view.scene().grid_lines, len(self.view.scene().lines.lines)):
            for j in range(len(self.view.scene().load_centre)):
                if self.view.scene().lines.lines[li].coordinates[-1] == [self.view.scene().load_centre[j][1],
                   self.view.scene().load_centre[j][2]]:
                    lins.append(menu.addAction(self.view.scene().lines.lines[li].name))
                    lins[-1].setIconVisibleInMenu(True)
        x = self.frameGeometry().left() + 50
        y = self.frameGeometry().y() + 50
        action = menu.exec_(QtCore.QPoint(x, y))
        if action != None:
            for li in range(len(self.view.scene().lines.lines)):
                if self.view.scene().lines.lines[li].name == action.text():
                    grid_path_len = self.view.traceGrid(None, coords=self.view.scene().lines.lines[li].coordinates)
            try:
                self.view.emit(SIGNAL('statusmsg'), 'Grid traced for %s (%s, %s Km)' % (action.text(),
                               self.view.scene().lines.lines[li].line_table, '{:0.1f}'.format(grid_path_len)))
            except:
                self.view.emit(SIGNAL('statusmsg'), 'No Grid line for %s' % action.text())

    def show_Ruler(self):
        comment = 'Scale Ruler Toggled'
        if self.view.scene().show_ruler:
            self.view.hide_Ruler()
            self.showRuler.setIcon(QtGui.QIcon('blank.png'))
            self.view.scene().show_ruler = False
            comment += ' Off'
        else:
            self.view.show_Ruler(self.view.scene().ruler, self.view.scene().ruler_ticks)
            self.showRuler.setIcon(QtGui.QIcon(self.check_icon))
            self.view.scene().show_ruler = True
            comment += ' On'
        self.view.emit(SIGNAL('statusmsg'), comment)

    def show_Resource(self):
   #      if self.resource_grid == '':
       #      return
        if self.resource is None:
            if self.sender().text()[:4] == 'Show':
                self.resource_year = self.base_year
            else:
                self.resource_year = str(self.sender().text())
            self.resource = Resource(self.resource_year, self.view.scene().resource_grid)
            self.resource.setWindowModality(QtCore.Qt.WindowModal)
            self.resource.setWindowFlags(self.resource.windowFlags() |
                              QtCore.Qt.WindowSystemMenuHint |
                              QtCore.Qt.WindowMinMaxButtonsHint)
            self.resource.procStart.connect(self.getResource)
            self.view.emit(SIGNAL('statusmsg'), 'Resource Window opened')
            self.resource.show()
       #      self.resource.exec_()

    @QtCore.pyqtSlot(str)
    def getResource(self, text):
        if text == 'goodbye':
          #   self.view.resourceGrid('hide')
            self.resource = None
            comment = 'Resource Window closed'
        elif text == 'hide':
            self.view.resourceGrid('hide')
            comment = 'Resource variable removed'
        elif text == 'show':
            self.view.resourceGrid('show', self.resource, year=self.resource_year)
            comment = 'Resource variable displayed. %s. Full Range: %s to %s; %s: %s to %s' % \
                      (self.view.resource_range[0], '{:0.1f}'.format(self.view.resource_range[2]),
                      '{:0.1f}'.format(self.view.resource_range[3]), self.view.resource_range[1],
                      '{:0.1f}'.format(self.view.resource_range[4]),
                      '{:0.1f}'.format(self.view.resource_range[5]))
        self.view.emit(SIGNAL('statusmsg'), comment)

    def show_FloatMenu(self):
        if self.floatmenu is None:
            self.floatmenu = FloatMenu(self.menuBar())
            self.floatmenu.setWindowModality(QtCore.Qt.WindowModal)
            self.floatmenu.setWindowFlags(self.floatmenu.windowFlags() |
                              QtCore.Qt.WindowSystemMenuHint |
                              QtCore.Qt.WindowMinMaxButtonsHint)
            self.floatmenu.procStart.connect(self.getFloatMenu)
            self.floatmenu.procAction.connect(self.getFloatAction)
            self.view.emit(SIGNAL('statusmsg'), 'Floating Menu opened')
            self.floatmenu.show()
      #       self.floatmenu.exec_()

    @QtCore.pyqtSlot(str)
    def getFloatMenu(self, text):
        if text == 'goodbye':
            self.floatmenu = None
            comment = 'Floating Menu closed'
        self.view.emit(SIGNAL('statusmsg'), comment)

    @QtCore.pyqtSlot(QtGui.QAction)
    def getFloatAction(self, action):
        action.trigger()

    def reshow_FloatMenu(self):
        if self.floatmenu is None:
            return
        self.floatmenu.exit()
        self.floatmenu = None
        self.show_FloatMenu()

    def show_FloatLegend(self, make_comment=True):
        if self.floatlegend is None:
            tech_data = {}
            flags = [self.view.scene().show_capacity, self.view.scene().show_generation,
                     self.view.scene().station_square, self.view.scene().show_fossil]
            for key in self.view.scene().areas.keys():
                tech_data[key] = [self.view.scene().areas[key], self.view.scene().colors[key], flags]
            self.floatlegend = FloatLegend(tech_data, self.view.scene()._stations.stations, flags)
            self.floatlegend.setWindowModality(QtCore.Qt.WindowModal)
            self.floatlegend.setWindowFlags(self.floatlegend.windowFlags() |
                              QtCore.Qt.WindowSystemMenuHint |
                              QtCore.Qt.WindowMinMaxButtonsHint)
            self.floatlegend.procStart.connect(self.getFloatLegend)
            if make_comment:
                self.view.emit(SIGNAL('statusmsg'), 'Floating Legend opened')
            self.floatlegend.show()
            self.activateWindow()
      #       self.floatmenu.exec_()

    @QtCore.pyqtSlot(str)
    def getFloatLegend(self, text):
        if text == 'goodbye':
            self.floatlegend = None
            comment = 'Floating Legend closed'
        self.view.emit(SIGNAL('statusmsg'), comment)

    def reshow_FloatLegend(self):
        if self.floatlegend is not None:
            self.floatlegend.exit()
            self.floatlegend = None
            self.show_FloatLegend(make_comment=False)
        if self.view.scene().show_legend:
            self.view.hide_Legend(self.view.legend_pos)
            self.view.show_Legend(self.view.legend_pos)

    def show_Legend(self):
        comment = 'Legend Toggled'
        if self.view.scene().show_legend:
            self.view.hide_Legend()
            self.showLegend.setIcon(QtGui.QIcon('blank.png'))
            self.view.scene().show_legend = False
            comment += ' Off'
        else:
            self.view.show_Legend()
            self.showLegend.setIcon(QtGui.QIcon(self.check_icon))
            self.view.scene().show_legend = True
            comment += ' On'
        self.view.emit(SIGNAL('statusmsg'), comment)

    def show_Name(self):
        comment = 'Station Names Toggled'
        if self.view.scene().show_station_name:
            self.showName.setIcon(QtGui.QIcon('blank.png'))
            self.view.scene().show_station_name = False
            self.view.scene()._nameGroup.setVisible(False)
            self.view.scene()._fnameGroup.setVisible(False)
            comment += ' Off'
        else:
            self.showName.setIcon(QtGui.QIcon(self.check_icon))
            self.view.scene().show_station_name = True
            self.view.scene()._nameGroup.setVisible(True)
            if self.view.scene().show_fossil:
                self.view.scene()._fnameGroup.setVisible(True)
            comment += ' On'
        self.view.emit(SIGNAL('statusmsg'), comment)

    def exit(self):
        if self.altered_stations:
            self.saveScenario()
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
            if self.credits is None:
                pass
            else:
                lines.append('credits_pos=%s,%s' % (str(self.credits.pos().x() + add), str(self.credits.pos().y() + add)))
                lines.append('credits_size=%s,%s' % (str(self.credits.size().width()), str(self.credits.size().height())))
                if self.resource is not None:
                    lines.append('resource_pos=%s,%s' % (str(self.resource.pos().x() + add), str(self.resource.pos().y() + add)))
                    lines.append('resource_size=%s,%s' % (str(self.resource.size().width()), str(self.resource.size().height())))
                if self.floatlegend is not None:
                    lines.append('legend_pos=%s,%s' % (str(self.floatlegend.pos().x() + add),
                                 str(self.floatlegend.pos().y() + add)))
                    lines.append('legend_show=True')
                    lines.append('legend_size=%s,%s' % (str(self.floatlegend.size().width()),
                                 str(self.floatlegend.size().height())))
                if self.floatmenu is None:
                    lines.append('menu_show=False')
                else:
                    lines.append('menu_pos=%s,%s' % (str(self.floatmenu.pos().x() + add), str(self.floatmenu.pos().y() + add)))
                    lines.append('menu_show=True')
                    lines.append('menu_size=%s,%s' % (str(self.floatmenu.size().width()), str(self.floatmenu.size().height())))
                if self.floatstatus is not None:
                    lines.append('log_pos=%s,%s' % (str(self.floatstatus.pos().x() + add),
                                 str(self.floatstatus.pos().y() + add)))
                    lines.append('log_size=%s,%s' % (str(self.floatstatus.size().width()),
                                 str(self.floatstatus.size().height())))
            updates['Windows'] = lines
            SaveIni(updates)
        if self.credits is not None:
            if self.credits.initial:  # rude way to close all windows and exit
                sys.exit()
            self.credits.exit()
        if self.resource is not None:
            self.resource.exit()
        if self.floatmenu is not None:
            self.floatmenu.exit()
        if self.floatlegend is not None:
            self.floatlegend.exit()
        if self.floatstatus is not None:
            self.floatstatus.exit()
     #    event.accept()

    def writeStations(self, scenario, description):
        the_scenario = scenario
        if scenario[-4:] == '.csv' or scenario[-4:] == '.xls' or scenario[-5:] == '.xlsx':
            pass
        else:
            the_scenario += '.xls'
        if os.path.exists(self.scenarios + the_scenario):
            if os.path.exists(self.scenarios + the_scenario + '~'):
                os.remove(self.scenarios + the_scenario + '~')
            os.rename(self.scenarios + the_scenario, self.scenarios + the_scenario + '~')
        ctr = 0
        d = 0
        fields = ['Station Name', 'Technology', 'Latitude', 'Longitude', 'Maximum Capacity (MW)',
                  'Turbine', 'Rotor Diam', 'No. turbines', 'Area']
        for stn in self.view.scene()._stations.stations:
            if stn.power_file is not None:
                if 'Power File' not in fields:
                    fields.append('Power File')
            if stn.storage_hours is not None:
                if 'Storage Hours' not in fields:
                    fields.append('Storage Hours')
            if stn.grid_line is not None:
                if 'Grid Line' not in fields:
                    fields.append('Grid Line')
            if stn.direction is not None:
                if 'Direction' not in fields:
                    fields.append('Direction')
        if the_scenario[-4:] == '.csv':
            upd_file = open(self.scenarios + the_scenario, 'wb')
            if description != '':
                upd_file.write('Description:,"' + description + '"\n')
            upd_writer = csv.writer(upd_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
            upd_writer.writerow(fields)
            for stn in self.view.scene()._stations.stations:
                if stn.scenario != 'Existing':
                    if stn.scenario == scenario:
                        new_line = []
                        new_line.append(stn.name)
                        new_line.append(stn.technology)
                        new_line.append(str(stn.lat))
                        new_line.append(str(stn.lon))
                        new_line.append(str(stn.capacity))
                        if stn.technology == 'Wind':
                            new_line.append(stn.turbine)
                            new_line.append(str(stn.rotor))
                            new_line.append(str(stn.no_turbines))
                        else:
                            new_line.append('')
                            new_line.append('0')
                            new_line.append('0')
                        new_line.append(str(stn.area))
                        if stn.power_file is not None:
                            new_line.append(str(stn.power_file))
                        if stn.grid_line is not None:
                            new_line.append('"' + str(stn.grid_line) + '"')
               # need to add fields for storage hours + fix
                        if stn.direction is not None:
                            new_line.append(str(stn.direction))
                        upd_writer.writerow(new_line)
                        ctr += 1
            upd_file.close()
        else:
            wb = xlwt.Workbook()
            fnt = xlwt.Font()
            fnt.bold = True
            styleb = xlwt.XFStyle()
            styleb.font = fnt
            lens = []
            for i in range(len(fields)):
                lens.append(len(fields[i]))
            ws = wb.add_sheet(scenario[:scenario.find('.')])
            d = 0
            if description != '':
                ws.write(ctr, 0, 'Description:')
                i = 0
                r = 0
                while i >= 0:
                    r += 1
                    i = description.find('\n', i + 1)
                ws.write_merge(ctr, ctr, 1, 9, description)
                if r > 1:
                    ws.row(0).height = int((fnt.height * 1.2)) * r
                d = -1
                ctr += 1
            for i in range(len(fields)):
                ws.write(ctr, i, fields[i])
            for stn in self.view.scene()._stations.stations:
                if stn.scenario != 'Existing':
                    if stn.scenario == scenario or stn.scenario == the_scenario:
                        ctr += 1
                        ws.write(ctr, 0, stn.name)
                        lens[0] = max(lens[0], len(stn.name))
                        ws.write(ctr, 1, stn.technology)
                        lens[1] = max(lens[1], len(stn.technology))
                        ws.write(ctr, 2, stn.lat)
                        lens[2] = max(lens[2], len(str(stn.lat)))
                        ws.write(ctr, 3, stn.lon)
                        lens[3] = max(lens[3], len(str(stn.lon)))
                        ws.write(ctr, 4, stn.capacity)
                        lens[4] = max(lens[4], len(str(stn.capacity)))
                        if stn.technology == 'Wind':
                            ws.write(ctr, 5, stn.turbine)
                            lens[5] = max(lens[5], len(stn.turbine))
                            ws.write(ctr, 6, stn.rotor)
                            lens[6] = max(lens[6], len(str(stn.rotor)))
                            ws.write(ctr, 7, stn.no_turbines)
                            lens[7] = max(lens[7], len(str(stn.no_turbines)))
                        else:
                            ws.write(ctr, 6, 0)
                            ws.write(ctr, 7, 0)
                        ws.write(ctr, 8, stn.area)
                        lens[8] = max(lens[8], len(str(stn.area)))
                        if stn.power_file is not None:
                            ws.write(ctr, fields.index('Power File'), str(stn.power_file))
                            lens[fields.index('Power File')] = max(lens[fields.index('Power File')],
                                                               len(str(stn.power_file)))
                        if stn.grid_line is not None:
                            ws.write(ctr, fields.index('Grid Line'), str(stn.grid_line))
                            lens[fields.index('Grid Line')] = max(lens[fields.index('Grid Line')],
                                                               len(str(stn.grid_line)))
                        if stn.storage_hours is not None:
                            if stn.storage_hours != self.view.scene().tshours:
                                ws.write(ctr, fields.index('Storage Hours'), stn.storage_hours)
                                lens[fields.index('Storage Hours')] = max(lens[fields.index('Storage Hours')],
                                                                      len(str(stn.storage_hours)))
                        if stn.direction is not None:
                            ws.write(ctr, fields.index('Direction'), str(stn.direction))
                            lens[fields.index('Direction')] = max(lens[fields.index('Direction')],
                                                               len(str(stn.direction)))
            for c in range(9):
                if lens[c] * 275 > ws.col(c).width:
                    ws.col(c).width = lens[c] * 275
            ws.set_panes_frozen(True)  # frozen headings instead of split panes
            ws.set_horz_split_pos(1 - d)  # in general, freeze after last heading row
            ws.set_remove_splits(True)  # if user does unfreeze, don't leave a split there
            wb.save(self.scenarios + the_scenario)
        if self.floatstatus and self.log_status:
            self.floatstatus.emit(SIGNAL('log'), '%s. Saved %s station(s) to %s' %
                            (str(QtCore.QDateTime.toString(QtCore.QDateTime.currentDateTime(),
                            'hh:mm:ss')), str(ctr + d), the_scenario))


def main():
    app = QtGui.QApplication(sys.argv)
    scene = WAScene()
    mw = MainWindow(scene)
    QtGui.QShortcut(QtGui.QKeySequence('q'), mw, mw.close)
    QtGui.QShortcut(QtGui.QKeySequence('x'), mw, mw.exit)
    ver = fileVersion()
    mw.setWindowTitle('SIREN (' + ver + ' - Beta) - ' + scene.model_name)
    scene_ratio = float(mw.view.scene().width()) / mw.view.scene().height()
    screen = QtGui.QDesktopWidget().availableGeometry()
    screen_ratio = float(screen.width()) / screen.height()
    if sys.platform == 'win32' or sys.platform == 'cygwin':
        pct = 0.85
    else:
        pct = 0.90
    if screen_ratio > scene_ratio:
        h = int(screen.height() * pct)
        w = int(screen.height() * pct * scene_ratio)
    else:
        w = int(screen.width() * pct)
        h = int(screen.width() * pct / scene_ratio)
    mw.resize(w, h)
    rescale = w / mw.view.scene().width()
    mw.view.scale(rescale, rescale)
    mw.center()
    show_credits = True
    show_floatlegend = False
    show_floatmenu = False
    if mw.restorewindows:
        config = ConfigParser.RawConfigParser()
        config.read(mw.config_file)
        try:
            rw = config.get('Windows', 'main_size').split(',')
            if w != int(rw[0]) or h != int(rw[1]):
                mw.resize(int(rw[0]), int(rw[1]))
            mp = config.get('Windows', 'main_pos').split(',')
            mw.move(int(mp[0]), int(mp[1]))
            vw = config.get('Windows', 'main_view').split(',')
            vw = map(float, vw)
            tgt_width = vw[2] - vw[0]
            tgt_height = vw[3] - vw[1]
            mw.view.centerOn(vw[0] + (vw[2] - vw[0]) / 2, vw[1] + (vw[3] - vw[1]) / 2)
            cur_width = mw.view.mapToScene(mw.view.width(), mw.view.height()).x() - mw.view.mapToScene(0, 0).x()
            cur_height = mw.view.mapToScene(mw.view.width(), mw.view.height()).y() - mw.view.mapToScene(0, 0).y()
            ctr = 0
            while (cur_width > tgt_width or cur_height > tgt_height) and ctr < 30:
                mw.view.zoomIn()
                cur_width = mw.view.mapToScene(mw.view.width(), mw.view.height()).x() - mw.view.mapToScene(0, 0).x()
                cur_height = mw.view.mapToScene(mw.view.width(), mw.view.height()).y() - mw.view.mapToScene(0, 0).y()
                ctr += 1
            mw.view.width(), mw.view.height()
        except:
            pass
        try:
            cr = config.get('Windows', 'credits_show')
            if cr.lower() in ['false', 'no', 'off']:
                show_credits = False
        except:
            pass
        try:
            cr = config.get('Windows', 'legend_show')
            if cr.lower() in ['true', 'yes', 'on']:
                show_floatlegend = True
        except:
            pass
        try:
            cr = config.get('Windows', 'menu_show')
            if cr.lower() in ['true', 'yes', 'on']:
                show_floatmenu = True
        except:
            pass
    if show_credits:
        mw.showCredits(initial=True)
    mw.show()
    if show_floatmenu:
        mw.show_FloatMenu()
    if show_floatlegend:
        mw.show_FloatLegend()
  #   mw.activateWindow()
    app.exec_()
    app.deleteLater()
    sys.exit()


if '__main__' == __name__:
    main()
