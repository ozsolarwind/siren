#!/usr/bin/python
#
#  Copyright (C) 2015 Sustainable Energy Now Inc., Angus King
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
