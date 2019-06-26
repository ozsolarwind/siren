#!/usr/bin/python3
#
#  Copyright (C) 2015-2019 Sustainable Energy Now Inc., Angus King
#
#  getparents.py - This file is part of SIREN.
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
def getParents(aparents):
    parents = []
    for key, value in aparents:
        for key2, value2 in aparents:
            if key2 == key:
                continue
            value = value.replace(key2, value2)
        for key2, value2 in parents:
            if key2 == key:
                continue
            value = value.replace(key2, value2)
        parents.append((key, value))
    return parents
