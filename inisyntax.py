#!/usr/bin/python3
#
#  Copyright (C) 2015-2022 Sustainable Energy Now Inc., Angus King
#
#  inisyntax.py - This file is part of SIREN.
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

from PyQt5.QtCore import QRegExp
from PyQt5.QtGui import QColor, QTextCharFormat, QFont, QSyntaxHighlighter


def format(color, style=''):
    """Return a QTextCharFormat with the given attributes.
    """
    _color = QColor()
    _color.setNamedColor(color)

    _format = QTextCharFormat()
    _format.setForeground(_color)
    if 'bold' in style:
        _format.setFontWeight(QFont.Bold)
    if 'italic' in style:
        _format.setFontItalic(True)

    return _format


# Syntax styles that can be shared by all languages
STYLES = {
    'operator': format('darkGreen'),
    'group': format('darkMagenta'),
    'comment': format('darkBlue'),
    'self': format('black'),
}


class IniHighlighter (QSyntaxHighlighter):
    """Syntax highlighter for the .ini files.
    """
    comment = ['^\;', '^#']
    group = ['^\[']
    operators = ['\=']

    def __init__(self, document, line=None):
        QSyntaxHighlighter.__init__(self, document)
        self.line = line
        rules = []
        rules += [(r'%s' % c, 0, STYLES['comment'])
            for c in self.comment]
        rules += [(r'%s' % g, 0, STYLES['group'])
            for g in self.group]
        rules += [(r'%s' % o, 0, STYLES['operator'])
           for o in self.operators]
         # Build a QRegExp for each pattern
        self.rules = [(QRegExp(pat), index, fmt)
            for (pat, index, fmt) in rules]

    def highlightBlock(self, text):
        """Apply syntax highlighting to the given block of text.
        """
        # Do other syntax formatting
        for expression, nth, format in self.rules:
            index = expression.indexIn(text, 0)
            if index == 0:
                self.setFormat(index, len(text), format)
                break
            if index > 0:
                self.setFormat(0, index, format)
                break
        if self.currentBlock().blockNumber() == self.line:
            line = self.currentBlock().blockNumber()
            fmt = QTextCharFormat()
            fmt.setBackground(QColor("yellow"))
            self.setFormat(0, len(text), fmt)
        self.setCurrentBlockState(0)
