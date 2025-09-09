#!/usr/bin/python3
#
#  Copyright (C) 2019-2025 Angus King
#
#  zoompan.py - This file is used by SIREN.
#
#  This is free software: you can redistribute it and/or modify
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
# Based on: https://gist.github.com/tacaswell/3144287
# and https://stackoverflow.com/questions/10374930/matplotlib-annotating-a-3d-scatter-plot
from math import ceil, sqrt
import matplotlib
if matplotlib.__version__ > '3.5.1':
    matplotlib.use('Qt5Agg')
else:
    matplotlib.use('TkAgg')
from mpl_toolkits.mplot3d import proj3d
from matplotlib import __version__ as matplotlib_version
#from matplotlib.lines import Line2D
#from matplotlib.collections import PathCollection
from mpl_toolkits.mplot3d.art3d import Path3DCollection
import warnings


class ZoomPanX():
    def __init__(self, yformat=None):
        self.base_xlim = None
        self.base_ylim = None
        self.base_zlim = None
        self.xlabel = None
        self.axis = 'x'
        self.angle = 0
        self.roll = 0
        self.elev = 30
        self.azim = -60
        self.d3 = False
        self.datapoint = None
        self.msg = None
        self.tbar = None
        self.cur_xlim = None
        self.press = None
        self.month = None
        self.week = None
        self.keys = ''
        self.yformat = yformat
        self.flex_ticks = False
        self.flex_on = False
        self.step = 168
        self.pick = True
        self.hide = False
        self.days = 3

    def zoom_pan(self, ax, base_scale=2., annotate=False, dropone=False, flex_ticks=False, pick=True, mth_labels=None, days=3):
        def set_flex():
            if self.flex_ticks:
                cur_xlim = ax.get_xlim()
                if cur_xlim[1] - cur_xlim[0] > self.step:
                    if not self.flex_on:
                        ax.set_xticks(self.x_ticks_s)
                        ax.set_xticklabels(self.x_labels_s)
                        ax.set_xlim(cur_xlim)
                        self.flex_on = True
                else:
                    if self.flex_on:
                        ax.set_xticks(self.x_ticks)
                        ax.set_xticklabels(self.x_labels)
                        ax.set_xlim(cur_xlim)
                        self.flex_on = False

        def zoom(event):
            if event.inaxes != ax:
                return
            if self.base_xlim is None:
                self.base_xlim = ax.get_xlim()
                self.base_ylim = ax.get_ylim()
                try:
                    self.base_zlim = ax.get_zlim()
                    self.d3 = True
                except:
                    self.d3 = False
            # get the current x and y limits
            cur_xlim = ax.get_xlim()
            cur_ylim = ax.get_ylim()
            if self.d3:
                cur_zlim = ax.get_zlim()
            xdata = event.xdata # get event x location
            ydata = event.ydata # get event y location
            if event.button == 'up':
                # deal with zoom in
                scale_factor = 1 / base_scale
            elif event.button == 'down':
                # deal with zoom out
                scale_factor = base_scale
            else:
                # deal with something that should never happen
                scale_factor = 1
                print('(109)', event.button)
            # set new limits
            if self.d3:
                z_left = ydata - cur_zlim[0]
                z_right = cur_zlim[1] - ydata
            if self.axis == 'x':
                ax.set_xlim([xdata - (xdata - cur_xlim[0]) * scale_factor,
                            xdata + (cur_xlim[1] - xdata) * scale_factor])
            elif self.axis == 'y':
                ax.set_ylim([ydata - (ydata - cur_ylim[0]) * scale_factor,
                            ydata + (cur_ylim[1] - ydata) * scale_factor])
            elif self.axis == 'z':
                ax.set_zlim([ydata - (ydata - cur_zlim[0]) * scale_factor,
                            ydata + (cur_zlim[1] - ydata) * scale_factor])
            set_flex()
       #     ax.figure.canvas.draw() # force re-draw
            ax.figure.canvas.draw_idle() # force re-draw

        def get_xyz_mouse_click(event, ax):
            """
            Get coordinates clicked by user
            """
            if ax.M is None:
                return {}
            xd, yd = event.xdata, event.ydata
            p = (xd, yd)
            edges = ax.tunit_edges()
            ldists = [(proj3d._line2d_seg_dist(p0, p1, p), i) for \
                        i, (p0, p1) in enumerate(edges)]
            ldists.sort()
            # nearest edge
            edgei = ldists[0][1]
            p0, p1 = edges[edgei]
            # scale the z value to match
            x0, y0, z0 = p0
            x1, y1, z1 = p1
            d0 = sqrt(pow(x0 - xd, 2) + pow(y0 - yd, 2))
            d1 = sqrt(pow(x1 - xd, 2) + pow(y1 - yd, 2))
            dt = d0 + d1
            z = d1 / dt * z0 + d0 / dt * z1
            x, y, z = proj3d.inv_transform(xd, yd, z, ax.M)
            return x, y, z

        def onPress(event):
            if self.base_xlim is None:
                self.base_xlim = ax.get_xlim()
                self.base_ylim = ax.get_ylim()
                try:
                    self.base_zlim = ax.get_zlim()
                    self.d3 = True
                except:
                    self.d3 = False

            if self.d3 and matplotlib_version > '3.0.3':
                try:
                    x, y, z = get_xyz_mouse_click(event, ax)
                except:
                    return
                if not self.pick:
                    return
                self.datapoint = [[-1, x, y, z]]
                self.msg = '{}: {:.2f}\n{}: {:.2f}\n{}: {:.2f}'.format(
                            ax.get_xlabel(), self.datapoint[0][1], ax.get_ylabel(),
                            self.datapoint[0][2], ax.get_zlabel(), self.datapoint[0][3])
                # If we have previously displayed another label, remove it first
                if hasattr(ax, 'label'):
                    try:
                        ax.label.remove()
                    except:
                        pass
                x2, y2, zs = proj3d.inv_transform(self.datapoint[0][1], self.datapoint[0][2],
                             self.datapoint[0][3], ax.get_proj())
                ax.label = ax.annotate(self.msg, xy = (x2, y2), xytext = (0, 20),
                           textcoords = 'offset points', ha = 'right', va = 'bottom',
                           bbox = dict(boxstyle = 'round,pad=0.5', alpha = 0.5),
                           zorder=100,
                           arrowprops = dict(arrowstyle = '->',
                                             connectionstyle = 'arc3,rad=0'))
                return
          #  if self.tbar._active is not None:
           #     return
            if event.button == 3: # reset?
                self.month = None
                self.week = None
                if self.base_xlim is not None:
                    ax.set_xlim(self.base_xlim)
                    set_flex()
                    ax.figure.canvas.draw()
                    return
            if event.inaxes != ax:
                return
            if self.axis == 'x':
                self.cur_xlim = ax.get_xlim()
                self.press = event.xdata
            elif self.axis == 'y':
                self.cur_ylim = ax.get_ylim()
                self.press = event.ydata
            elif self.axis == 'z':
                self.cur_zlim = ax.get_zlim()
                self.press = event.zdata

        def onRelease(event):
            self.press = None
            if not self.pick:
                return
            if self.datapoint is not None:
                # If we have previously displayed another label, remove it first
                if hasattr(ax, 'label'):
                    try:
                        ax.label.remove()
                    except:
                        pass
                x2, y2, _ = proj3d.proj_transform(self.datapoint[0][1], self.datapoint[0][2],
                            self.datapoint[0][3], ax.get_proj())
                ax.label = ax.annotate(self.msg, xy = (x2, y2), xytext = (0, 20),
                           textcoords = 'offset points', ha = 'right', va = 'bottom',
                           bbox = dict(boxstyle = 'round,pad=0.5', alpha = 0.5),
                           zorder=100,
                           arrowprops = dict(arrowstyle = '->',
                                             connectionstyle = 'arc3,rad=0'))
                set_flex()
            ax.figure.canvas.draw()

        def onMotion(event):
            if self.base_xlim is None:
                self.base_xlim = ax.get_xlim()
                self.base_ylim = ax.get_ylim()
                try:
                    self.base_zlim = ax.get_zlim()
                    self.d3 = True
                except:
                    self.d3 = False
            if self.press is None:
                return
            if event.inaxes != ax:
                return
            if self.axis == 'x':
                dx = event.xdata - self.press
                self.cur_xlim -= dx
                ax.set_xlim(self.cur_xlim)
            elif self.axis == 'y':
                dy = event.ydata - self.press
                self.cur_ylim -= dy
                ax.set_ylim(self.cur_ylim)
            elif self.axis == 'z':
                dz = event.ydata - self.press
                self.cur_zlim -= dz
                ax.set_zlim(self.cur_zlim)
            ax.figure.canvas.draw()

        def onKey(event):
            if event.key is None:
                return
            try:
                event_key = event.key.lower()
            except:
                event_key = event.key
            if event_key == 'r': # reset
                self.keys = ''
                self.month = None
                ax.figure.canvas.manager.set_window_title(self.wtitle)
                if mth_labels is not None:
                    ax.set_ylabel(self.ylabel)
                self.week = None
                if self.base_xlim is not None:
                    ax.set_xlim(self.base_xlim)
                    ax.set_ylim(self.base_ylim)
                    if self.d3:
                        if hasattr(ax, 'label'):
                            try:
                                ax.label.remove()
                            except:
                                pass
                            self.datapoint = None
                        ax.set_zlim(self.base_zlim)
                    set_flex()
                    if self.d3:
                        ax.view_init(elev=self.elev, azim=self.azim)
                    ax.figure.canvas.draw()
                    return
            if event_key == 'pageup':
                if self.axis != 'x':
                    return
                if self.week is not None:
                    self.week -= 1
                    if self.month is None:
                        if self.week < 0:
                            self.week = ceil(self.base_xlim[1] / self.step) - 1
                        strt = self.week * self.step # 24 * 7 hours or 48 * 7 half-hours per week
                    else:
                        if self.week < 0:
                            if self.month == 1 and self.mth_xlim[2] == 1416:
                                self.week = 3
                            else:
                                self.week = 4
                        strt = self.mth_xlim[self.month] + self.week * self.step
                    ax.set_xlim([strt, strt + self.step])
                else:
                    if self.month is None or self.month == 0:
                        self.month = len(self.mth_xlim) - 2
                    else:
                        self.month -= 1
                    ax.set_xlim([self.mth_xlim[self.month], self.mth_xlim[self.month + 1]])
                ax.figure.canvas.draw()
            elif event_key == 'pagedown':
                if self.axis != 'x':
                    return
                if self.week is not None:
                    self.week += 1
                    if self.month is None:
                        if self.week >= ceil(self.base_xlim[1] / self.step):
                            self.week = 0
                        strt = self.week * self.step
                    else:
                        if self.week >= 5 or \
                          (self.month == 1 and self.week >= 4 and self.mth_xlim[2] == 1416):
                            self.week = 0
                        strt = self.mth_xlim[self.month] + self.week * self.step
                    ax.set_xlim([strt, strt + self.step])
                else:
                    if self.month is None:
                        self.month = 0
                    else:
                        if self.month >= len(self.mth_xlim) - 2:
                            self.month = 0
                        else:
                            self.month += 1
                    ax.set_xlim([self.mth_xlim[self.month], self.mth_xlim[self.month + 1]])
                ax.figure.canvas.draw()
            elif event_key == 'd':
                if self.axis != 'x':
                    return
                if self.keys == 'd':
                    incr1 = self.step / 7
                else:
                    incr1 = 0
                incr2 = incr1 + (self.step / 7) * self.days
                self.keys = 'd'
                xlim = ax.get_xlim()
                ax.set_xlim(xlim[0] + incr1, xlim[0] + incr2)
                set_flex()
                ax.figure.canvas.draw()
            elif event_key == 'h': # hide legend
                font = self.fig_axes[-1].get_legend().prop
                handles = []
                labels = []
                for axs in self.fig_axes:
                    handle, label = axs.get_legend_handles_labels()
                    handles.extend(handle)
                    labels.extend(label)
                if self.hide:
                    us = ''
                    self.hide = False
                else:
                    us = '_'
                    self.hide = True
                for l in range(len(labels)):
                    labels[l] = us + labels[l]
                # reverse the order
                self.fig_axes[-1].legend(labels=labels)
                # reverse the order
                self.fig_axes[-1].legend(handles[::-1], labels[::-1], prop=font).set_draggable(True)
                if self.yformat is not None:
                    ax.yaxis.set_major_formatter(self.yformat)
                ax.figure.canvas.draw()
            elif event_key == 'l':
                self.keys = 'l'
                font = self.fig_axes[-1].get_legend().prop
                handles = []
                labels = []
                for axs in self.fig_axes:
                    handle, label = axs.get_legend_handles_labels()
                    handles.extend(handle)
                    labels.extend(label)
                # reverse the order
                self.fig_axes[-1].legend(handles[::-1], labels[::-1], prop=font).set_draggable(True)
                if self.yformat is not None:
                    ax.yaxis.set_major_formatter(self.yformat)
                ax.figure.canvas.draw()
           #     if self.yformat is not None:
             #       ax.yaxis.set_major_formatter(self.yformat)
            elif event_key == 'm':
                if self.axis != 'x':
                    return
                self.keys = 'm'
                self.week = None
                if self.month is None:
                    self.month = 0
                elif self.month >= len(self.mth_xlim) - 2:
                    self.month = 0
                else:
                    self.month += 1
                ax.figure.canvas.manager.set_window_title(self.wtitle + f'-{self.month+1:02}')
                if mth_labels is not None:
                    ax.set_ylabel(mth_labels[self.month] + '\n' + self.ylabel)
                ax.set_xlim([self.mth_xlim[self.month], self.mth_xlim[self.month + 1]])
                set_flex()
                ax.figure.canvas.draw()
            elif event_key == 'w':
                if self.axis != 'x':
                    return
                self.keys = 'w'
                if self.week is None:
                    self.week = 0
                else:
                    self.week += 1
                if self.month is None:
                    if self.week >= ceil(self.base_xlim[1] / self.step):
                        self.week = 0
                    strt = self.week * self.step # 24 * 7 hours per week
                else:
                    if self.week >= 5 or \
                      (self.month == 1 and self.week >= 4 and self.mth_xlim[2] == 1416):
                        self.week = 0
                    strt = self.mth_xlim[self.month] + self.week * self.step
                ax.set_xlim([strt, strt + self.step])
                set_flex()
                ax.figure.canvas.draw()
            elif event_key == 't':
                if self.axis != 'x':
                    return
                self.keys = 't'
                if self.week is None:
                    self.week = 0
                else:
                    self.week += 1
                if self.month is None:
                    if self.week >= ceil(self.base_xlim[1] / self.step):
                        self.week = 0
                    strt = self.week * self.step # 24 * 7 hours per week
                else:
                    if self.week >= 5 or \
                      (self.month == 1 and self.week >= 4 and self.mth_xlim[2] == 1416):
                        self.week = 0
                    strt = self.mth_xlim[self.month] + self.week * self.step
                ax.set_xlim([strt, strt + self.step * 2])
                set_flex()
                ax.figure.canvas.draw()
            elif event.key >= '0' and event.key <= '9' and len(self.keys) > 0:
                if self.axis != 'x':
                    return
                if self.keys == 'd':
                    self.days = int(event.key)
                    incr1 = 0
                    incr2 = incr1 + (self.step / 7) * self.days
                    xlim = ax.get_xlim()
                    ax.set_xlim(xlim[0] + incr1, xlim[0] + incr2)
                    set_flex()
                    ax.figure.canvas.draw()
                    return
                elif self.keys[-1] == 'l':
                    if event.key == '0':
                        ncol =1
                    else:
                        ncol = int(event.key)
                    font = self.fig_axes[-1].get_legend().prop
                    handles = []
                    labels = []
                    for axs in self.fig_axes:
                        handle, label = axs.get_legend_handles_labels()
                        handles.extend(handle)
                        labels.extend(label)
                    self.fig_axes[-1].legend(handles[::-1], labels[::-1], prop=font, ncol=ncol).set_draggable(True)
                    if self.yformat is not None:
                        ax.yaxis.set_major_formatter(self.yformat)
                    ax.figure.canvas.draw()
                    return
                elif self.keys[-2:] == 'm1':
                    self.keys = ''
                    if event.key < '3':
                        self.month = 10 + int(event.key) - 1
                        ax.figure.canvas.manager.set_window_title(self.wtitle + f'-{self.month+1:02}')
                        if mth_labels is not None:
                            ax.set_ylabel(mth_labels[self.month] + '\n' + self.ylabel)
                    else:
                        return
                elif self.keys[-1:] == 'm':
                    if event.key == '0':
                        self.month = 0
                    else:
                        self.month = int(event.key) - 1
                        ax.figure.canvas.manager.set_window_title(self.wtitle + f'-{self.month+1:02}')
                        if mth_labels is not None:
                            ax.set_ylabel(mth_labels[self.month] + '\n' + self.ylabel)
                    self.keys += event.key
                elif self.keys[-1] == 'w' or (len(self.keys) > 1 and self.keys[-2] == 'w'):
                    if self.keys[-1] == 'w':
                        wk = event_key
                    else:
                        wk = self.keys[-1] + event_key
                    if wk == '0':
                        self.week = 0
                    else:
                        self.week = int(wk) - 1
                    if self.month is None:
                        if self.week >= ceil(self.base_xlim[1] / self.step):
                            self.week = 0
                        strt = self.week * self.step # 24 * 7 hours per week
                    else:
                        if self.week >= 5 or \
                          (self.month == 1 and self.week >= 4 and self.mth_xlim[2] == 1416):
                            self.week = 0
                        strt = self.mth_xlim[self.month] + self.week * self.step
                    self.keys += event.key
                    ax.set_xlim([strt, strt + self.step])
                    set_flex()
                    ax.figure.canvas.draw()
                    return
                elif self.keys[-1] == 't' or (len(self.keys) > 1 and self.keys[-2] == 't'):
                    if self.keys[-1] == 't':
                        wk = event_key
                    else:
                        wk = self.keys[-1] + event_key
                    if wk == '0':
                        self.week = 0
                    else:
                        self.week = int(wk) - 1
                    if self.month is None:
                        if self.week >= ceil(self.base_xlim[1] / self.step):
                            self.week = 0
                        strt = self.week * self.step
                    else:
                        if self.week >= 5 or \
                          (self.month == 1 and self.week >= 4 and self.mth_xlim[2] == 1416):
                            self.week = 0
                        strt = self.mth_xlim[self.month] + self.week * self.step
                    self.keys += event.key
                    ax.set_xlim([strt, strt + self.step * 2])
                    set_flex()
                    ax.figure.canvas.draw()
                    return
                ax.set_xlim([self.mth_xlim[self.month], self.mth_xlim[self.month + 1]])
                set_flex()
                ax.figure.canvas.draw()
            elif event_key == 'x':
                if self.d3:
                    if self.axis == 'x':
                        if self.angle <= 0:
                            self.angle = 90
                        else:
                            self.angle = -90
                    try:
                        ax.view_init(elev=self.angle, azim=-self.angle, roll=0, vertical_axis='x')
                    except:
                        try:
                            ax.view_init(elev=self.angle, azim=-self.angle, vertical_axis='x')
                        except:
                            ax.view_init(elev=self.angle, azim=-self.angle)
                    ax.figure.canvas.draw()
                elif self.axis != 'x':
                    self.press = None
                self.axis = 'x'
            elif event_key == 'y':
                if self.d3:
                    if self.axis == 'y':
                        if self.angle <= 0:
                            self.angle = 90
                        else:
                            self.angle = -90
                    try:
                        ax.view_init(elev=self.angle, azim=-self.angle, roll=0, vertical_axis='y')
                    except:
                        try:
                            ax.view_init(elev=self.angle, azim=-self.angle, vertical_axis='y')
                        except:
                            ax.view_init(elev=self.angle, azim=-self.angle)
                    ax.figure.canvas.draw()
                elif self.axis != 'y':
                    self.press = None
                self.axis = 'y'
            elif event_key == 'z':
                if self.d3:
                    if self.axis == 'z':
                        if self.angle <= 0:
                            self.angle = 90
                        else:
                            self.angle = -90
                    try:
                        ax.view_init(elev=self.angle, azim=-self.angle, roll=0, vertical_axis='z')
                    except:
                        try:
                            ax.view_init(elev=self.angle, azim=-self.angle, vertical_axis='z')
                        except:
                            ax.view_init(elev=self.angle, azim=-self.angle)
                    ax.figure.canvas.draw()
                elif self.axis != 'z':
                    self.press = None
                self.axis = 'z'
            elif self.d3 and (event.key == '>' or event.key == '.'):
                self.roll = self.roll + 90
                if self.roll > 270:
                    self.roll = 0
                try:
                    ax.view_init(elev=self.angle, azim=-self.angle, roll=self.roll, vertical_axis=self.axis)
                    ax.figure.canvas.draw()
                except:
                    pass
            elif event.key == 'right':
                xlim = ax.get_xlim()
                if xlim[1] >= self.base_xlim[1]:
                    ax.set_xlim([0, xlim[1] - xlim[0]])
                else:
                    ax.set_xlim([xlim[1], xlim[1] + (xlim[1] - xlim[0])])
                ax.figure.canvas.draw()
            elif event.key == 'left':
                xlim = ax.get_xlim()
                if xlim[0] <= self.base_xlim[0]:
                    ax.set_xlim([self.base_xlim[1] - (xlim[1] - xlim[0]), self.base_xlim[1]])
                else:
                    ax.set_xlim([xlim[0] - (xlim[1] - xlim[0]), xlim[0]])
                ax.figure.canvas.draw()
            elif event.key == 'up':
                ylim = ax.get_ylim()
                if ylim[1] >= self.base_ylim[1]:
                    ax.set_ylim([0, ylim[1] - ylim[0]])
                else:
                    ax.set_ylim([ylim[1], ylim[1] + (ylim[1] - ylim[0])])
                ax.figure.canvas.draw()
            elif event.key == 'down':
                ylim = ax.get_ylim()
                if ylim[0] <= self.base_ylim[0]:
                    ax.set_ylim([self.base_ylim[1] - (ylim[1] - ylim[0]), self.base_ylim[1]])
                else:
                    if ylim[0] - (ylim[1] - ylim[0]) < self.base_ylim[0]:
                        ax.set_ylim([self.base_ylim[0], self.base_ylim[0] + (ylim[1] - ylim[0])])
                    else:
                        ax.set_ylim([ylim[0] - (ylim[1] - ylim[0]), ylim[0]])
                ax.figure.canvas.draw()

        def onPick(event):
            if not isinstance(event.artist, Path3DCollection): # just 3D picking for the moment
                return
            if matplotlib_version > '3.0.3':
                return
            if not self.pick:
                return
            self.datapoint = None
            if len(event.ind) > 0:
                self.datapoint = []
                if self.d3:
                    x, y, z = event.artist._offsets3d # 2021-05-21
                    for n in event.ind:
                        self.datapoint.append([n, event.artist._offsets3d[0][n],
                            event.artist._offsets3d[1][n], event.artist._offsets3d[2][n]])
                    self.msg = '{:d}: {}: {:.2f}\n{}: {:.2f}\n{}: {:.2f}'.format(self.datapoint[0][0],
                          ax.get_xlabel(), self.datapoint[0][1], ax.get_ylabel(),
                          self.datapoint[0][2], ax.get_zlabel(), self.datapoint[0][3])
                    # If we have previously displayed another label, remove it first
                    if hasattr(ax, 'label'):
                        try:
                            ax.label.remove()
                        except:
                            pass
                    x2, y2, zs = proj3d.inv_transform(self.datapoint[0][1], self.datapoint[0][2],
                                self.datapoint[0][3], ax.get_proj())
                    ax.label = ax.annotate(self.msg, xy = (x2, y2), xytext = (0, 20),
                                textcoords = 'offset points', ha = 'right', va = 'bottom',
                                bbox = dict(boxstyle = 'round,pad=0.5', alpha = 0.5),
                                arrowprops = dict(arrowstyle = '->',
                                                  connectionstyle = 'arc3,rad=0'))
                set_flex()
                ax.figure.canvas.draw()

        warnings.filterwarnings('ignore', module ='.*zoompan.*') # ignore warnings include legend hide
        the_days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
        if mth_labels is not None:
            self.ylabel = ax.get_ylabel()
            self.window = ax.figure.canvas.manager.get_window_title()
        self.pick = pick
        self.xlabel = ax.get_xlabel()
        self.wtitle = ax.figure.canvas.manager.get_window_title()
        self.base_xlim = ax.get_xlim() # remember x base
        if self.base_xlim[1] > 8784:
            self.step = 168 * 2 # 48 * 7 half-hours per week
        else:
            self.step = 168 # 24 * 7 hours per week
        self.base_ylim = ax.get_ylim() # remember y base
        self.flex_ticks = flex_ticks
        self.days = days
        if self.flex_ticks:
            self.x_ticks = []
            self.x_labels = []
            self.x_ticks_s = []
            self.x_labels_s = []
            l = 0
            for xtick in ax.get_xticks():
                self.x_ticks.append(xtick)
                if l % 7 == 0:
                    self.x_ticks_s.append(xtick)
                l += 1
            l = 0
            for xtick in ax.get_xticklabels():
                self.x_labels.append(xtick.get_text())
                if l % 7 == 0:
                    self.x_labels_s.append(xtick.get_text())
                l += 1
            self.flex_on = False
            set_flex()
        try:
            self.base_zlim = ax.get_zlim() # remember z base for 3D
            self.d3 = True
            self.elev = ax.elev
            self.azim = ax.azim
        except:
            self.d3 = False
        if self.base_xlim[1] == 8784 or self.base_xlim[1] == 8784 * 2: # leap year
            the_days[1] = 29
        x = 0
        self.mth_xlim = [x]
        if self.base_xlim[1] > 8784:
            mult = 48
        else:
            mult = 24
        if dropone:
            no_of_days = 364
        else:
            no_of_days = 365
        if self.base_xlim[1] == no_of_days or self.base_xlim[1] == no_of_days + 1: # daily data
            if self.base_xlim[1] == no_of_days + 1: # leap year
                the_days[1] = 29
            mult = 1
        for days in the_days:
            x += days * mult
            self.mth_xlim.append(x)
            if x > self.base_xlim[1]:
                break
        fig = ax.get_figure() # get the figure of interest
        self.fig_axes = fig.axes # save axes in case 2nd y axis
        self.tbar = fig.canvas.toolbar # get toolbar
        # attach the call back
        # 'axis_enter_event'
        # 'axis_leave_event'
        fig.canvas.mpl_connect('button_press_event', onPress)
        fig.canvas.mpl_connect('button_release_event', onRelease)
        # 'draw_event'
        # 'figure_enter_event'
        # 'figure_leave_event'
        fig.canvas.mpl_connect('key_press_event', onKey)
        # 'key_release_event'
        fig.canvas.mpl_connect('motion_notify_event', onMotion)
        fig.canvas.mpl_connect('pick_event', onPick)
        # 'resize_event'
        fig.canvas.mpl_connect('scroll_event', zoom)
        #return the function
        return zoom
