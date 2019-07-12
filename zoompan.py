#!/usr/bin/python3
#
#  Copyright (C) 2019 Angus King
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

the_days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]


class ZoomPanX():
    def __init__(self):
        self.base_xlim = None
        self.tbar = None
        self.cur_xlim = None
        self.xpress = None
        self.month = None
        self.week = None
        self.keys = ''

    def zoom_pan(self, ax, base_scale=2.):
        def zoom(event):
            if event.inaxes != ax:
                return
            if self.base_xlim is None:
                self.base_xlim = ax.get_xlim()
            # get the current x and y limits
            cur_xlim = ax.get_xlim()
            cur_ylim = ax.get_ylim()
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
                print('(56)', event.button)
            # set new limits
            # Get distance from the cursor to the edge of the figure frame
            x_left = xdata - cur_xlim[0]
            x_right = cur_xlim[1] - xdata
            y_top = ydata - cur_ylim[0]
            y_bottom = cur_ylim[1] - ydata
            ax.set_xlim([xdata - x_left*scale_factor,
                        xdata + x_right*scale_factor])
            ax.figure.canvas.draw() # force re-draw

        def onPress(event):
            if self.base_xlim is None:
                self.base_xlim = ax.get_xlim()
            if self.tbar._active is not None:
                return
            if event.button == 3: # reset?
                self.month = None
                self.week = None
                if self.base_xlim is not None:
                    ax.set_xlim(self.base_xlim)
                    ax.figure.canvas.draw()
                    return
            if event.inaxes != ax:
                return
            self.cur_xlim = ax.get_xlim()
            self.xpress = event.xdata

        def onRelease(event):
            self.xpress = None
            ax.figure.canvas.draw()

        def onMotion(event):
            if self.base_xlim is None:
                self.base_xlim = ax.get_xlim()
            if self.xpress is None:
                return
            if event.inaxes != ax:
                return
            dx = event.xdata - self.xpress
            self.cur_xlim -= dx
            ax.set_xlim(self.cur_xlim)
            ax.figure.canvas.draw()

        def onKey(event):
            if event.key.lower() == 'r':
                self.keys = ''
                self.month = None
                self.week = None
                if self.base_xlim is not None:
                    ax.set_xlim(self.base_xlim)
                    ax.figure.canvas.draw()
                    return
            if event.key == 'pageup':
                if self.week is not None:
                    self.week -= 1
                    if self.month is None:
                        if self.week < 0:
                            self.week = 52
                        strt = self.week * 168 # 24 * 7 hours per week
                    else:
                        if self.week < 0:
                            if self.month == 1 and self.mth_xlim[2] == 1416:
                                self.week = 3
                            else:
                                self.week = 4
                        strt = self.mth_xlim[self.month] + self.week * 168
                    ax.set_xlim([strt, strt + 168])
                else:
                    if self.month is None or self.month == 0:
                        self.month = 11
                    else:
                        self.month -= 1
                    ax.set_xlim([self.mth_xlim[self.month], self.mth_xlim[self.month + 1]])
                ax.figure.canvas.draw()
            elif event.key == 'pagedown':
                if self.week is not None:
                    self.week += 1
                    if self.month is None:
                        if self.week >= 52:
                            self.week = 0
                        strt = self.week * 168 # 24 * 7 hours per week
                    else:
                        if self.week >= 5 or \
                          (self.month == 1 and self.week >= 4 and self.mth_xlim[2] == 1416):
                            self.week = 0
                        strt = self.mth_xlim[self.month] + self.week * 168
                    ax.set_xlim([strt, strt + 168])
                else:
                    if self.month is None:
                        self.month = 0
                    else:
                        if self.month >= 11:
                            self.month = 0
                        else:
                            self.month += 1
                    ax.set_xlim([self.mth_xlim[self.month], self.mth_xlim[self.month + 1]])
                ax.figure.canvas.draw()
            elif event.key.lower() == 'm':
                self.keys = 'm'
                self.week = None
                if self.month is None:
                    self.month = 0
                elif self.month >= 11:
                    self.month = 0
                else:
                    self.month += 1
                ax.set_xlim([self.mth_xlim[self.month], self.mth_xlim[self.month + 1]])
                ax.figure.canvas.draw()
            elif event.key.lower() == 'w':
                self.keys = ''
                if self.week is None:
                    self.week = 0
                else:
                    self.week += 1
                if self.month is None:
                    if self.week >= 52:
                        self.week = 0
                    strt = self.week * 168 # 24 * 7 hours per week
                else:
                    if self.week >= 5 or \
                      (self.month == 1 and self.week >= 4 and self.mth_xlim[2] == 1416):
                        self.week = 0
                    strt = self.mth_xlim[self.month] + self.week * 168
                ax.set_xlim([strt, strt + 168])
                ax.figure.canvas.draw()
            elif event.key >= '0' and event.key <= '9':
                if self.keys[-2:] == 'm1':
                    self.keys = ''
                    if event.key < '3':
                        self.month = 10 + int(event.key) - 1
                    else:
                        return
                elif self.keys[-1:] == 'm':
                    if event.key == '0':
                        self.month = 0
                    else:
                        self.month = int(event.key) - 1
                    self.keys += event.key
                ax.set_xlim([self.mth_xlim[self.month], self.mth_xlim[self.month + 1]])
                ax.figure.canvas.draw()
        the_days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
        self.base_xlim = ax.get_xlim() # remember x base
        if self.base_xlim[1] == 8784: # leap year
            the_days[1] = 29
        x = 0
        self.mth_xlim = [x]
        for days in the_days:
            x += days * 24
            self.mth_xlim.append(x)
        fig = ax.get_figure() # get the figure of interest
        self.tbar = fig.canvas.toolbar # get toolbar
        # attach the call back
        fig.canvas.mpl_connect('scroll_event', zoom)
        fig.canvas.mpl_connect('button_press_event', onPress)
        fig.canvas.mpl_connect('button_release_event', onRelease)
        fig.canvas.mpl_connect('motion_notify_event', onMotion)
        fig.canvas.mpl_connect('key_press_event', onKey)
        #return the function
        return zoom
