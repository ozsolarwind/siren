#!/usr/bin/python3
#
#  Copyright (C) 2015-2019 Sustainable Energy Now Inc., Angus King
#
#  senuser.py - This file is part of SIREN.
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

#
# return current userid

import os
try:
    import pwd
except:
    pass
import sys


def getUser():
    if sys.platform == 'win32' or sys.platform == 'cygwin':   # windows
        return os.environ.get("USERNAME")

    elif sys.platform == 'darwin':   # osx64
        return os.environ.get("USERNAME")

    elif sys.platform == 'linux2':   # linux
        return pwd.getpwuid(os.geteuid()).pw_name
    else:
        return os.environ.get("USERNAME")

def techClean(tech, full=False):
    cleantech = tech.replace('_', ' ').title()
    cleantech = cleantech.replace('Cst', 'CST')
    cleantech = cleantech.replace('Pv', 'PV')
    if full:
        alll = [['Cf', 'CF'], ['hi ', 'HI '], ['Lcoe', 'LCOE'], ['Mw', 'MW'],
                ['ni ', 'NI '], ['Npv', 'NPV'], ['Tco2E', 'tCO2e']]
        for each in alll:
            cleantech = cleantech.replace(each[0], each[1])
    return cleantech
