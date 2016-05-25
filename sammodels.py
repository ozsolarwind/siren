#!/usr/bin/python
#
#  Copyright (C) 2015 Sustainable Energy Now Inc., Angus King
#
#  sammodels.py - This file is part of SIREN.
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
#  The routines in this program have been derived from models
#  developed by The National Renewable Energy Laboratory (NREL)
#  Center for Renewable Energy Resources. Copyright for the
#  routines remain with them.
#
#  getDNI is derived from the DISC DNI Model
#  <http://rredc.nrel.gov/solar/models/DISC/>
#
#  getDHI is derived from the NREL DNI-GHI to DHI Calculator
#  <https://sam.nrel.gov/sites/sam.nrel.gov/files/content/documents/xls/DNI-GHI_to_DHI_Calculator.xlsx>
#

from math import *


def getDNI(ghi=0, hour=0, lat=0, lon=0, press=1013.25, zone=8, debug=False):
    hour_of_year = hour
    day_of_year = int((hour_of_year - 1) / 24) + 1
    latitude = lat
    longitude = lon
    day_angle = 6.283185 * (day_of_year - 1) / 365.
    time_zone = zone
    pressure = press
    etr = 1370 * (1.00011 + 0.034221 * cos(day_angle) + 0.00128 * sin(day_angle) +
          0.000719 * cos(2 * day_angle) + 0.000077 * sin(2 * day_angle))
    dec = (0.006918 - 0.399912 * cos(day_angle) + 0.070257 * sin(day_angle) -
          0.006758 * cos(2 * day_angle) + 0.000907 * sin(2 * day_angle) -
          0.002697 * cos(3 * day_angle) + 0.00148 * sin(3 * day_angle)) * (180. / 3.14159)
    eqt = (0.000075 + 0.001868 * cos(day_angle) - 0.032077 * sin(day_angle) -
          0.014615 * cos(2 * day_angle) - 0.040849 * sin(2 * day_angle)) * (229.18)
    hour_angle = 15 * (hour_of_year - 12 - 0.5 + eqt / 60 + ((longitude - time_zone * 15) * 4) / 60)
    zenith_angle = acos(cos(radians(dec)) * cos(radians(latitude)) * cos(radians(hour_angle)) +
          sin(radians(dec)) * sin(radians(latitude))) * (180. / 3.14159)
    if zenith_angle < 80:
        am = 1 / (cos(radians(zenith_angle)) + 0.15 / pow(93.885 - zenith_angle, 1.253)) \
        * (pressure / 1013.25)
    else:
        am = 0.
    if am > 0:
        kt = ghi / (cos(radians(zenith_angle)) * etr)
    else:
        kt = 0.
    a = 0.
    if kt > 0:
        if kt > 0.6:
            a = -5.743 + 21.77 * kt - 27.49 * pow(kt, 2) + 11.56 * pow(kt, 3)
        elif kt < 0.6:
            a = 0.512 - 1.56 * kt + 2.286 * pow(kt, 2) - 2.222 * pow(kt, 3)
    b = 0.
    if kt > 0:
        if kt > 0.6:
            b = 41.4 - 118.5 * kt + 66.05 * pow(kt, 2) + 31.9 * pow(kt, 3)
        elif kt < 0.6:
            b = 0.37 + 0.962 * kt
    c = 0.
    if kt > 0:
        if kt > 0.6:
            c = -47.01 + 184.2 * kt - 222 * pow(kt, 2) + 73.81 * pow(kt, 3)
        elif kt < 0.6:
            c = -0.28 + 0.932 * kt - 2.048 * pow(kt, 2)
    kn = 0.
    if kt > 0:
        kn = a + b * exp(c * am)
    knc = 0.
    if kt > 0:
        knc = 0.886 - 0.122 * am + 0.0121 * pow(am, 2) - 0.000653 * pow(am, 3) + \
              0.000014 * pow(am, 4)
    if debug:
        print 'hour', hour
        print 'latitude', latitude
        print 'longitude', longitude
        print 'time_zone', time_zone
        print 'pressure', pressure
        print 'day_angle', day_angle
        print 'etr', etr
        print 'dec', dec
        print 'eqt', eqt
        print 'hour_angle', hour_angle
        print 'zenith_angle', zenith_angle
        print 'am', am
        print 'kt', kt
        print 'a', a
        print 'b', b
        print 'c', c
        print 'kn', kn
        print 'knc', knc
    if kt > 0:
        if etr * (knc - kn) >= 0:
            return etr * (knc - kn)
    return 0.


def getDHI(ghi=0, dni=0, hour=0, lat=0, azimuth=0., tilt=0., reflectance=0.2, debug=False):
# Ic = Ib cos(i) + Idh cos^2(B/2) + p Ih sin^2(B/2)
# Ic = insolation on a collector; Ib = beam insolation; Idh = diffuse insolation on horizontal; Ih = total insolation on horizontal
# i = incidence angle; B = surface tilt; p = gnd reflection
# Idh = Ih - Ibh; Ibh = Ib sin(a)  where a = altitude angle

    hour_of_year = hour   # B
    day_of_year = int((hour_of_year - 1) / 24) + 1   # H
    declination_angle = asin(sin(-23.45 * pi / 180.) * cos(360. / 365. * (10.5 + day_of_year) * pi / 180.)) * 180. / pi
    latitude = lat
# I
    if hour_of_year % 24 != 0:
        hour_angle = (15 * (12 - hour_of_year % 24)) + 7.5
    else:
        hour_angle = -172.5
# L
    sun_rise_hour_angle = acos(-1 * tan(latitude * pi / 180.) * tan(declination_angle * pi / 180.)) * 180. / pi
# J
    if abs(hour_angle) - 7.5 < sun_rise_hour_angle and abs(hour_angle) + 7.5 > sun_rise_hour_angle:
        hour_angle_2 = abs(hour_angle) - 7.5 + (1 - (((abs(hour_angle) + 7.5) - sun_rise_hour_angle) / 15)) * 7.5
#                         1          1        12    345   6          6      5             4     32      1
    else:
        hour_angle_2 = 15 * (12 - hour_of_year % 24) + 7.5
#                     1     2    3
# K
    if hour_angle < 0:
        sun_rise_set_adjusted_hour_angle = abs(hour_angle_2) * -1
    else:
        sun_rise_set_adjusted_hour_angle = hour_angle_2
# M
    sun_rise_hr_am = 12 - (acos(-1 * tan(latitude * pi / 180.) * tan(declination_angle * pi / 180.)) * 180. / pi) / 15.
# N
    sun_set_hr_pm = acos(-1 * tan(latitude * pi / 180.) * tan(declination_angle * pi / 180.)) * 180. / pi / 15. + 12
# O
    if (acos(-tan(latitude * pi / 180.) * tan(declination_angle * pi / 180.)) * 180. / pi) > abs(sun_rise_set_adjusted_hour_angle):
        sun_rise_set_hsr = 'UP'
    else:
        sun_rise_set_hsr = 'DOWN'
# Q
    altitude_angle = asin((sin(latitude * pi / 180.) * sin(declination_angle * pi / 180.)) +
                     (cos(latitude * pi / 180.) * cos(declination_angle * pi / 180.) *
                     cos(sun_rise_set_adjusted_hour_angle * pi / 180.))) * 180. / pi
# R
    if sun_rise_set_adjusted_hour_angle > 0:
        solar_azimuth = abs(acos(((cos(declination_angle * pi / 180.) * sin(latitude * pi / 180.) *
                        cos(sun_rise_set_adjusted_hour_angle * pi / 180.)) -
                        (sin(declination_angle * pi / 180.) * cos(latitude * pi / 180.))) /
                        cos(altitude_angle * pi / 180.)) * 180. / pi)
    else:
        solar_azimuth = -1 * abs(acos(((cos(declination_angle * pi / 180.) * sin(latitude * pi / 180.) *
                        cos(sun_rise_set_adjusted_hour_angle * pi / 180.)) -
                        (sin(declination_angle * pi / 180.) * cos(latitude * pi / 180.))) /
                        cos(altitude_angle * pi / 180.)) * 180. / pi)
# S
    incidence_angle = acos((cos(altitude_angle * pi / 180.) *
        cos((solar_azimuth - azimuth) * pi / 180.) * sin(tilt * pi / 180.)) +
        ((sin(altitude_angle * pi / 180.) * cos(tilt * pi / 180.)))) * 180. / pi
# T
    if incidence_angle < 90:
        beam_component = dni * cos(incidence_angle * pi / 180.)
    else:
        beam_component = 0.
# U, V
    if altitude_angle > 0:
        diffuse_component = (ghi - (dni * sin(altitude_angle * pi / 180.))) * \
            pow(cos((tilt / 2) * pi / 180.), 2)
        reflected_component = reflectance * ghi * pow(sin((tilt / 2) * pi / 180.), 2)
    else:
        diffuse_component = 0.
        reflected_component = 0.
# W
    total_wh = beam_component + diffuse_component + reflected_component
# X
    check_total = total_wh - ghi
# Y
    if diffuse_component < 0:
        dhi_negative = 0.
    else:
        dhi_negative = diffuse_component
# Z
    dhi_rounded = round(dhi_negative, 1)
    if debug:
        print 'ghi', ghi
        print 'dni', dni
        print 'latitude', latitude
        print 'hour_of_year', hour_of_year
        print 'day_of_year', day_of_year
        print 'declination_angle', declination_angle
        print 'hour_angle', hour_angle
        print 'sun_rise_hour_angle', sun_rise_hour_angle
        print 'hour_angle_2', hour_angle_2
        print 'sun_rise_set_adjusted_hour_angle', sun_rise_set_adjusted_hour_angle
        print 'sun_rise_hr_am', sun_rise_hr_am
        print 'sun_set_hr_pm', sun_set_hr_pm
        print 'sun_rise_set_hsr', sun_rise_set_hsr
        print 'altitude_angle', altitude_angle
        print 'solar_azimuth', solar_azimuth
        print 'incidence_angle', incidence_angle
        print 'beam_component', beam_component
        print 'diffuse_component', diffuse_component
        print 'reflected_component', reflected_component
        print 'total_wh', total_wh
        print 'check_total', check_total
        print 'dhi_negative', dhi_negative
    return dhi_rounded


def getZenith(hour=0, lat=0, lon=0, zone=8, debug=False):
    hour_of_year = hour
    day_of_year = int((hour_of_year - 1) / 24) + 1
    latitude = lat
    longitude = lon
    day_angle = 6.283185 * (day_of_year - 1) / 365.
    time_zone = zone
    etr = 1370 * (1.00011 + 0.034221 * cos(day_angle) + 0.00128 * sin(day_angle) +
          0.000719 * cos(2 * day_angle) + 0.000077 * sin(2 * day_angle))
    dec = (0.006918 - 0.399912 * cos(day_angle) + 0.070257 * sin(day_angle) -
          0.006758 * cos(2 * day_angle) + 0.000907 * sin(2 * day_angle) -
          0.002697 * cos(3 * day_angle) + 0.00148 * sin(3 * day_angle)) * (180. / 3.14159)
    eqt = (0.000075 + 0.001868 * cos(day_angle) - 0.032077 * sin(day_angle) -
          0.014615 * cos(2 * day_angle) - 0.040849 * sin(2 * day_angle)) * (229.18)
    hour_angle = 15 * (hour_of_year - 12 - 0.5 + eqt / 60 + ((longitude - time_zone * 15) * 4) / 60)
    zenith_angle = acos(cos(radians(dec)) * cos(radians(latitude)) * cos(radians(hour_angle)) +
          sin(radians(dec)) * sin(radians(latitude))) * (180. / 3.14159)
    if debug:
        print 'hour', hour
        print 'latitude', latitude
        print 'longitude', longitude
        print 'time_zone', time_zone
        print 'day_angle', day_angle
        print 'etr', etr
        print 'dec', dec
        print 'eqt', eqt
        print 'hour_angle', hour_angle
        print 'zenith_angle', zenith_angle
    return zenith_angle
