#!/usr/bin/env python3
#
#  Copyright (C) 2024-2025 Angus King
#
#  senplot3d.py - This file is possibly part of SIREN.
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

from datetime import datetime
import plotly.graph_objects as go
import plotly

alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
mth_labels = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

def plotly_start():
    lines = ['<!DOCTYPE html>',
             '<head>',
             '<style>',
             '.fullscreen {',
             '  position: absolute;',
             '  top: 0;',
             '  left: 0;',
             '  bottom: 0;',
             '  right: 0;',
             '  overflow: auto;',
             '}',
             '</style>',
             '</head>',
             '<body>',
             '<div class="fullscreen" id="myDiv">',
             '<script src="https://cdn.plot.ly/plotly-2.32.0.min.js" charset="utf-8"></script>',
             '<script type="text/javascript">',
             'window.PlotlyConfig = {MathJaxConfig: "local"};']
    return lines

def plotly_end(datas):
    dats = ''
    for d in range(datas):
        dats += f'data{d}, '
    lines = [f'Plotly.newPlot("myDiv", [{dats[:-2]}], layout, config);',
             '</script>',
             '</div>',
             '</body>']
    return lines

def dict_to_str(dict_in):
    str_dict = str(dict_in)
    bits = str_dict.split("': ")
    for b in range(len(bits)):
        bits[b] = bits[b].replace("{'", "{")
        bits[b] = bits[b].replace('True', 'true')
        bits[b] = bits[b].replace('False', 'false')
        i = bits[b].rfind(", '")
        if i > 0:
            bits[b] = bits[b][:i + 2] + bits[b][i + 3:]
    str_dict = ''
 #   print(bits)
    for bit in bits:
   #     print(len(bit), bit)
        for c in range(len(bit) -1, -1, -1):
            if bit[c] == ' ':
                if bit[c - 1] == ',':
                    bit = bit[:c] + '\n' + bit[c + 1:]
                break
            bit = bit.replace('], [', '],\n    [')
        str_dict += bit + ': '
    return str_dict[:-2]


class PowerPlot3D():

    def replace_words(self, what, src, tgt):
        words = {'m': ['$MTH$', '$MONTH$'],
                 'y': ['$YEAR$'],
                 's': ['$SHEET$']}
        tgt_str = src
        if tgt == 'find':
            for wrd in words[what]:
                tgt_num = tgt_str.find(wrd)
                if tgt_num >= 0:
                    return tgt_num
                tgt_num = tgt_str.find(wrd.lower())
                if tgt_num >= 0:
                    return tgt_num
            return -1
        else:
            for wrd in words[what]:
                tgt_str = tgt_str.replace(wrd, tgt)
                tgt_str = tgt_str.replace(wrd.lower(), tgt)
        return tgt_str

    def __init__(self, in_colours, in_cperiod, in_interval, in_order, in_period, in_rows,
                 in_seasons, the_days, in_title, in_toprow, ws, in_year, in_zone_row,
                 months=3, html=None, background=False, contours=False, aspectmode='auto'):
        the_rows = [in_toprow[1] + 1]
        for m in range(len(the_days)):
            the_rows.append(the_rows[-1] + the_days[m] * in_interval)
        per1 = in_period
        per2 = in_cperiod
        if per1 in ['<none>', 'Year']: # full year of hourly figures
            strt_row = [in_toprow[1]]
            todo_rows = [in_rows]
        else:
            strt_row = []
            todo_rows = []
            if per1 in in_seasons.keys():
                strt_row = []
                todo_rows = []
                for s in in_seasons[per1]:
                    m = 0
                    strt_row.append(0)
                    while m < s:
                        strt_row[-1] = strt_row[-1] + the_days[m] * in_interval
                        m += 1
                    strt_row[-1] = strt_row[-1] + in_toprow[1]
                    todo_rows.append(the_days[s] * in_interval)
            else:
                i = mth_labels.index(per1)
                todo_mths = [i]
                if per2 != '<none>':
                    j = mth_labels.index(per2)
                    if j == i:
                        pass
                    elif j > i:
                        for k in range(i + 1, j + 1):
                            todo_mths.append(k)
                    else:
                        for k in range(i + 1, 12):
                            todo_mths.append(k)
                        for k in range(j + 1):
                            todo_mths.append(k)
                for s in todo_mths:
                    m = 0
                    strt_row.append(0)
                    while m < s:
                        strt_row[-1] = strt_row[-1] + the_days[m] * in_interval
                        m += 1
                    strt_row[-1] = strt_row[-1] + in_toprow[1]
                    todo_rows.append(the_days[s] * in_interval)# find columns
        tot_todo = 0
        for t in todo_rows:
            tot_todo += t
        if tot_todo / in_interval <= months * 31:
            do_months = False
        else:
            do_months = True
        variables = {}
        z = []
        for c in range(len(in_order) -1, -1, -1):
            col = in_order[c]
            for c2 in range(2, ws.ncols):
                try:
                    column = ws.cell_value(in_toprow[0], c2).replace('\n',' ')
                except:
                    column = str(ws.cell_value(in_toprow[0], c2))
                if in_zone_row > 0 and ws.cell_value(in_zone_row, c2) != '' and ws.cell_value(in_zone_row, c2) is not None:
                    column = ws.cell_value(in_zone_row, c2).replace('\n',' ') + '.' + column
                if column == col:
                    variables[col] = [c2, len(z)]
                    z.append([])
        # get data
        y_tickvals = []
        y_ticktext = []
        if do_months:
            zc = []
            for c in range(len(z)):
                zc.append([])
            for s in range(len(strt_row)):
                for p in range(strt_row[s], strt_row[s] + todo_rows[s], in_interval):
                    if p >= ws.nrows:
                        break
                    try:
                        if ws.cell_value(p + 1, 1)[8:16] == '01 00:00':
                            for c in range(len(z)):
                                z[c].append([])
                                zc[c].append([])
                                for i in range(in_interval):
                                    z[c][-1].append(0)
                                    zc[c][-1].append(0)
                    except:
                        break
                    c = 0
                    for row in range(p, p + in_interval):
                        if row + 1 >= ws.nrows:
                            break
                        try:
                            m = the_rows.index(row + 1)
                            y_tickvals.append(len(z[0]))
                            y_ticktext.append(mth_labels[m])
                        except:
                            pass
                        for key, values in variables.items():
                            try:
                                z[values[1]][-1][c] =  z[values[1]][-1][c] + ws.cell_value(row + 1, values[0])
                                zc[values[1]][-1][c] += 1
                            except:
                                break
                        c += 1
            for i in range(len(z)):
                for j in range(len(z[i])):
                    for k in range(len(z[i][j])):
                        try:
                            z[i][j][k] = z[i][j][k] / zc[i][j][k]
                        except:
                            pass
        else:
            for s in range(len(strt_row)):
                for p in range(strt_row[s], strt_row[s] + todo_rows[s], in_interval):
                    if p >= ws.nrows:
                        break
                    for c in range(len(z)):
                        z[c].append([])
                    for row in range(p, p + in_interval):
                        if row + 1 >= ws.nrows:
                            break
                        try:
                            m = the_rows.index(row + 1)
                            y_tickvals.append(len(z[0]))
                            y_ticktext.append(mth_labels[m])
                        except:
                            if ws.cell_value(row + 1, 1)[11:13] == '00':
                                try:
                                    day = datetime.strptime(ws.cell_value(row + 1, 1)[:10], '%Y-%m-%d').weekday()
                                    if day == 0:
                                        y_tickvals.append(len(z[0]))
                                        y_ticktext.append(ws.cell_value(row + 1, 1)[8:10].lstrip('0'))
                                except:
                                    pass
                        for key, values in variables.items():
                            try:
                                z[values[1]][-1].append(ws.cell_value(row + 1, values[0]))
                            except:
                                break
        data = []
        x = []
        for i in range(in_interval):
            x.append(i)
        periods = []
        for h in range(24):
            periods.append(f'{h:02}:00')
            if in_interval == 48:
                periods.append(f'{h:02}:30')
        x_tickvals = []
        x_ticktext = []
        step = int(len(periods) / 6)
        for t in range(step, len(periods), step):
            x_tickvals.append(t)
            x_ticktext.append(periods[t])
        y= []
        for i in range(len(z[0])):
            y.append(i + 1)
        for key, values in variables.items():
            colorscale = [[0, f'{in_colours[key.lower()]}'], [0.5, f'{in_colours[key.lower()]}'],
                          [1, f'{in_colours[key.lower()]}']]
            data.append({'x': x,
                         'y': y,
                         'z': z[values[1]],
                         'name': key,
                         'showlegend': True,
                         'type': 'surface',
                         'colorscale': colorscale,
                         'showscale': False
                        })
            if contours:
                data[-1]['contours'] = {'z': {'show': True}}
        titl = self.replace_words('y', in_title, str(in_year))
        if per1 in ['<none>', 'Year']:
            titl = self.replace_words('m', titl, '')
        elif in_cperiod == '<none>':
            titl = self.replace_words('m', titl, in_period)
        else:
            titl = self.replace_words('m', titl, in_period + ' to ' + in_cperiod)
        titl = titl.replace('  ', ' ')
        titl = titl.replace('Diurnal ', '')
        titl = titl.replace('Diurnal', '')
        titl = self.replace_words('s', titl, ws.name)
        layout = {'title': {'text': titl,
                            'font': {'size': 25},
                            'x': 0.5,
                            'xanchor': 'center',
                            'yanchor': 'top'
                           },
                  'scene': {'xaxis': {'type': 'category',
                                      'title': 'Hour',
                                      'tickmode': 'array',
                                      'tickvals': x_tickvals,
                                      'ticktext': x_ticktext,
                                      'autorange': 'reversed'
                                     },
                            'yaxis': {'type': 'category',
                                      'title': 'Period',
                                      'tickmode': 'array',
                                      'tickvals': y_tickvals,
                                      'ticktext': y_ticktext,
                                      'autorange': 'reversed'
                                     },
                            'zaxis': {
                                      'title': 'Power/MWh)'
                                     },
                            'aspectmode': f'{aspectmode}'
                           },
                  'showlegend': True,
                  'legend': {'xanchor': 'left',
                             'x': 0.75,
                             'yanchor': 'bottom',
                             'y': 0.5
                            }
                 }
        config = {'responsive': True}
        if background:
            for axis in ['xaxis', 'yaxis', 'zaxis']:
                layout['scene'][axis]['showbackground'] = True
                layout['scene'][axis]['backgroundcolor'] = '#E5ECF6'
                layout['scene'][axis]['gridcolor'] = 'white'
                if axis == 'zaxis':
                    layout['scene'][axis]['zerolinecolor'] = 'black'
                else:
                    layout['scene'][axis]['zerolinecolor'] = 'white'
        if html is not None:
            afile = html
            out = open(afile, 'w')
            for line in plotly_start():
                out.write(line + '\n')
            for d in range(len(data)):
                out.write(f'var data{d} = {dict_to_str(data[d])}\n')
            out.write('var layout = ' + dict_to_str(layout) + '\n')
            out.write('var config = ' + dict_to_str(config) + '\n')
            for line in plotly_end(len(data)):
                out.write(line + '\n')
            out.close()
        fig = go.Figure(data)
        fig.update_layout(layout)
        fig.show(config=config)


class TablePlot3D():

    def __init__(self, colours, title, ws, x_offset, x_name, y_label, y_names, z_label, html=None,
                 background=False, contours=False, aspectmode='auto'):
        x_tickvals = []
        x_ticktext = []
        y_rows = []
        if x_offset.isdigit():
            print('not yet')
            x_col = 0
        else:
            x_col = alphabet.index(x_offset)
        for y in range(len(y_names)):
            y_rows.append([-1, -1])
        y = -1
        y_ticktext = []
        for row in range(ws.nrows):
            if ws.cell_value(row, x_col) == x_name or (isinstance(x_name, int) and x_name == row):
                for col in range(x_col + 1, ws.ncols):
                    if ws.cell_value(row, col) is None or ws.cell_value(row, col) == '':
                        break
                    x_tickvals.append(col)
                    x_ticktext.append(ws.cell_value(row, col))
            try:
                y = y_names.index(ws.cell_value(row, x_col))
                y_rows[y][0] = row + 1
                continue
            except:
                pass
            if ws.cell_value(row, x_col) is None or ws.cell_value(row, x_col) == '':
                if y >= 0:
                    y_rows[y][1] = row
                y = -1
            if y >= 0:
                y_tick = f'{ws.cell_value(row, x_col)}'
                if y_tick in y_ticktext:
                    pass
                else:
                    y_ticktext.append(y_tick)
        if y >= 0:
            y_rows[y][1] = ws.nrows
        y_ticks = []
        for t in range(len(y_ticktext)):
            y_ticks.append(t + 1)
        y_tickvals = sorted(y_ticks)
        z = []
        data = []
        for yr in range(len(y_names)):
            z.append([])
            for row in range(y_rows[yr][0], y_rows[yr][1]):
                z[-1].append([])
                for col in x_tickvals:
                    if ws.cell_value(row, col) is None:
                        z[-1][-1].append('')
                    else:
                        z[-1][-1].append(ws.cell_value(row, col))
            if len(z[-1]) == 0: # no data
                del z[-1]
                continue
            if isinstance(colours, dict):
                if isinstance(colours[y_names[yr].lower()], list):
                    e = float(len(colours[y_names[yr].lower()]) - 1)
                    colorscale = []
                    for i in range(len(colours[y_names[yr].lower()])):
                        colorscale.append([i / e, colours[y_names[yr].lower()][i]])
                else:
                    colorscale = [[0, f'{colours[y_names[yr].lower()]}'], [1, f'{colours[y_names[yr].lower()]}']]
            else:
                colorscale = colours[len(z)]
            data.append({'x': x_tickvals,
                         'y': y_tickvals,
                         'z': z[-1],
                         'name': y_names[yr],
                         'showlegend': True,
                         'type': 'surface',
                         'colorscale': colorscale,
                         'showscale': False})
            if contours:
                data[-1]['contours'] = {'z': {'show': True}}
        layout = {'title': {'text': title,
                            'font': {'size': 25},
                            'x': 0.5,
                            'xanchor': 'center',
                            'yanchor': 'top'
                           },
                  'scene': {'xaxis': {'type': 'category',
                                      'title': x_name,
                                      'tickmode': 'array',
                                      'tickvals': x_tickvals,
                                      'ticktext': x_ticktext,
                                      'autorange': 'reversed'
                                     },
                            'yaxis': {'type': 'category',
                                      'title': y_label,
                                      'tickmode': 'array',
                                      'tickvals': y_tickvals,
                                      'ticktext': y_ticktext,
                                      'autorange': 'reversed'
                                     },
                            'zaxis': {
                                      'title': z_label
                                     },
                            'aspectmode': f'{aspectmode}'
                           },
                  'showlegend': True,
                  'legend': {'xanchor': 'left',
                             'x': 0.75,
                             'yanchor': 'bottom',
                             'y': 0.5
                            }
                 }
        config = {'responsive': True}
        if background:
            for axis in ['xaxis', 'yaxis', 'zaxis']:
                layout['scene'][axis]['showbackground'] = True
                layout['scene'][axis]['backgroundcolor'] = '#E5ECF6'
                layout['scene'][axis]['gridcolor'] = 'white'
                if axis == 'zaxis':
                    layout['scene'][axis]['zerolinecolor'] = 'black'
                else:
                    layout['scene'][axis]['zerolinecolor'] = 'white'
        if html is not None:
            afile = html
            out = open(afile, 'w')
            for line in plotly_start():
                out.write(line + '\n')
            for d in range(len(data)):
                out.write(f'var data{d} = {dict_to_str(data[d])}\n')
            out.write('var layout = ' + dict_to_str(layout) + '\n')
            out.write('var config = ' + dict_to_str(config) + '\n')
            for line in plotly_end(len(data)):
                out.write(line + '\n')
            out.close()
        fig = go.Figure(data)
        fig.update_layout(layout)
        fig.show(config=config)
