#!/usr/bin/python
#
#  Copyright (C) 2015-2018 Sustainable Energy Now Inc., Angus King
#
#  djikstra_4.py - This file is part of SIREN.
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
#  This program has been derived from code developed by
#  http://www.bogotobogo.com Copyright for the original work
#  remains with them.
#  http://www.bogotobogo.com/python/files/Dijkstra/Dijkstra_shortest_path.py

import sys
import heapq
from math import sin, cos, asin, sqrt, radians

RADIUS = 6367.   # radius of earth in km


class Vertex:
    def __init__(self, node):
        self.id = node
        self.adjacent = {}
         # Set distance to infinity for all nodes
        self.distance = float(sys.maxint)
         # Mark all nodes unvisited
        self.visited = False
         # Predecessor
        self.previous = None

    def add_neighbor(self, neighbor, weight=0):
        self.adjacent[neighbor] = weight

    def get_connections(self):
        return self.adjacent.keys()

    def get_id(self):
        return self.id

    def get_weight(self, neighbor):
        return self.adjacent[neighbor]

    def set_distance(self, dist):
        self.distance = dist
        return

    def get_distance(self):
        return self.distance

    def set_previous(self, prev):
        self.previous = prev

    def set_visited(self):
        self.visited = True

#    def __str__(self):
#        return str(self.id) + ' adjacent: ' + str([x.id for x in self.adjacent])


class Graph:
    def __init__(self):
        self.vert_dict = {}
        self.num_vertices = 0

    def __iter__(self):
        return iter(self.vert_dict.values())

    def add_vertex(self, node):
        self.num_vertices = self.num_vertices + 1
        new_vertex = Vertex(node)
        self.vert_dict[node] = new_vertex
        return new_vertex

    def get_vertex(self, n):
        if n in self.vert_dict:
            return self.vert_dict[n]
        else:
            return None

    def add_edge(self, frm, to, cost=0):
        if frm not in self.vert_dict:
            self.add_vertex(frm)
        if to not in self.vert_dict:
            self.add_vertex(to)
        self.vert_dict[frm].add_neighbor(self.vert_dict[to], cost)
        self.vert_dict[to].add_neighbor(self.vert_dict[frm], cost)

    def get_vertices(self):
        return self.vert_dict.keys()

    def set_previous(self, current):
        self.previous = current

    def get_previous(self, current):
        return self.previous


def shortest(v, path):
    ''' make shortest path from v.previous'''
    if v.previous:
        path.append(v.previous.get_id())
        shortest(v.previous, path)
    return


def dijkstra(aGraph, start):
   #  print '''Dijkstra's shortest path'''
     # Set the distance for the start node to zero
    start.set_distance(0)

     # Put tuple pair into the priority queue
    unvisited_queue = [(v.get_distance(), v) for v in aGraph]
    heapq.heapify(unvisited_queue)

    while len(unvisited_queue):
         # Pops a vertex with the smallest distance
        uv = heapq.heappop(unvisited_queue)
        current = uv[1]
        current.set_visited()

        for next in current.adjacent:
             # if visited, skip
            if next.visited:
                continue
            new_dist = current.get_distance() + current.get_weight(next)

            if new_dist < next.get_distance():
                next.set_distance(new_dist)
                next.set_previous(current)
#                print 'updated : current = %s next = %s new_dist = %s' \
#                        %(current.get_id(), next.get_id(), str(next.get_distance()))
#            else:
#                print 'not updated : current = %s next = %s new_dist = %s' \
#                        %(current.get_id(), next.get_id(), str(next.get_distance()))

         # Rebuild heap
         # 1. Pop every item
        while len(unvisited_queue):
            heapq.heappop(unvisited_queue)
         # 2. Put all vertices not visited into the queue
        unvisited_queue = [(v.get_distance(), v) for v in aGraph if not v.visited]
        heapq.heapify(unvisited_queue)
    return


class Shortest:
    def isBetween(self, a, b, c):
        crossproduct = (c[1] - a[1]) * (b[0] - a[0]) - (c[0] - a[0]) * (b[1] - a[1])
        if abs(crossproduct) > 0.0001:  # sys.float_info.epsilon:
            return False    # (or != 0 if using integers)
        dotproduct = (c[0] - a[0]) * (b[0] - a[0]) + (c[1] - a[1]) * (b[1] - a[1])
        if dotproduct < 0:
            return False
        squaredlengthba = (b[0] - a[0]) * (b[0] - a[0]) + (b[1] - a[1]) * (b[1] - a[1])
        if dotproduct > squaredlengthba:
            return False
        return True

    def Distance(self, y1, x1, y2, x2):
# find the differences between the coordinates
        dy = y2 - y1
        dx = x2 - x1
        ra13 = pow(sin(dy / 2.), 2) + cos(y1) * cos(y2) * pow(sin(dx / 2.), 2)
        return 2 * asin(min(1, sqrt(ra13)))

    def actualDistance(self, y1d, x1d, y2d, x2d):
        x1 = radians(x1d)
        y1 = radians(y1d)
        x2 = radians(x2d)
        y2 = radians(y2d)
        dst = self.Distance(y1, x1, y2, x2)
        return round(abs(dst) * RADIUS, 2)

    def __init__(self, lines, source, target, grid):
        self.source = source
        self.target = target
        self.lines = lines
        self.grid = grid # existing grid lines count
        self.edges = {}
        self.g = Graph()
        for li in range(len(self.lines)):
            vert1 = str(self.lines[li].coordinates[0])
            for pt in range(1, len(self.lines[li].coordinates)):
                vert2 = str(self.lines[li].coordinates[pt])
                dist = self.actualDistance(self.lines[li].coordinates[pt][0], self.lines[li].coordinates[pt][1],
                       self.lines[li].coordinates[pt - 1][0], self.lines[li].coordinates[pt - 1][1])
                self.g.add_edge(vert1, vert2, dist)
                self.edges[vert1 + vert2] = li
                vert1 = vert2
        for li in range(self.grid, len(self.lines)):
            for l2 in range(len(self.lines)):
                if li == l2:
                    continue
                for pt in range(1, len(self.lines[l2].coordinates)):
                    if self.lines[l2].coordinates[pt - 1] == self.lines[li].coordinates[-1] or \
                      self.lines[l2].coordinates[pt] == self.lines[li].coordinates[-1]:
                        continue
                    if self.isBetween(self.lines[l2].coordinates[pt - 1], self.lines[l2].coordinates[pt],
                      self.lines[li].coordinates[-1]):
                        dist = self.actualDistance(self.lines[l2].coordinates[pt - 1][0], self.lines[l2].coordinates[pt - 1][1],
                               self.lines[li].coordinates[-1][0], self.lines[li].coordinates[-1][1])
                        vert1 = str(self.lines[l2].coordinates[pt - 1])
                        vert2 = str(self.lines[li].coordinates[-1])
                        self.g.add_edge(vert1, vert2, dist)
                        self.edges[vert1 + vert2] = l2
                        dist = self.actualDistance(self.lines[l2].coordinates[pt][0], self.lines[l2].coordinates[pt][1],
                               self.lines[li].coordinates[-1][0], self.lines[li].coordinates[-1][1])
                        vert1 = str(self.lines[l2].coordinates[pt])
                        vert2 = str(self.lines[li].coordinates[-1])
                        self.g.add_edge(vert1, vert2, dist)
                        self.edges[vert1 + vert2] = l2
                        break
     #    print 'Graph data:'
      #   print 'source', self.source, 'target', self.target
       #  for v in self.g:
        #     for w in v.get_connections():
         #        vid = v.get_id()
          #       wid = w.get_id()
           #      print '(%s ,%s, %s)'  % (vid, wid, str(v.get_weight(w)))
        if self.g.get_vertex(str(self.source)) is None:
            self.g.add_vertex(str(self.source))
        dijkstra(self.g, self.g.get_vertex(str(self.source)))
        target = self.g.get_vertex(str(self.target))
        try:
            self.path = [target.get_id()]
            shortest(target, self.path)
        except:
            self.path = []           

    def getPath(self):
        the_path = []
        for i in range(len(self.path)):
            it = self.path[i].replace('[', '')
            it = it.replace(']', '')
            it = it.replace(' ', '')
            bit = it.split(',')
            the_path.append([float(bit[0]), float(bit[1])])
        return the_path

    def getLines(self):
        the_lines = []
        for i in range(1, len(self.path)):
            try:
                the_lines.append(self.edges[self.path[i - 1] + self.path[i]])
            except:
                try:
                    the_lines.append(self.edges[self.path[i] + self.path[i - 1]])
                except:
                    pass
        the_lines = list(set(the_lines))
        return the_lines
