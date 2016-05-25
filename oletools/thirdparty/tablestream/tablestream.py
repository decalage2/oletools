#!/usr/bin/env python
"""
tablestream

tablestream can format table data for pretty printing as text,
to be displayed on the console or written to any file-like object.
The table data can be provided as rows, each row is an iterable of
cells. The text in each cell is wrapped to fit into a maximum width
set for each column.
Contrary to many table pretty printing libraries, TableStream writes
each row to the output as soon as it is provided, and the whole table
does not need to be built in memory before printing.
It is therefore suitable for large tables, or tables that take time to
be processed row by row.

Author: Philippe Lagadec - http://www.decalage.info
License: BSD, see source code or documentation
"""

#=== LICENSE ==================================================================

# tablestream is copyright (c) 2015-2016 Philippe Lagadec (http://www.decalage.info)
# All rights reserved.
#
# Redistribution and use in source and binary forms, with or without modification,
# are permitted provided that the following conditions are met:
#
#  * Redistributions of source code must retain the above copyright notice, this
#    list of conditions and the following disclaimer.
#  * Redistributions in binary form must reproduce the above copyright notice,
#    this list of conditions and the following disclaimer in the documentation
#    and/or other materials provided with the distribution.
#
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
# ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
# WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
# FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
# DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
# SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
# CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
# OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
# OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.


#------------------------------------------------------------------------------
# CHANGELOG:
# 2015-11-01 v0.01 PL: - first version
# 2016-01-01 v0.02 PL: - added styles, color support
# 2016-04-19 v0.03 PL: - enable colorclass on Windows, fixed issue #39
# 2016-05-25 v0.04 PL: - updated for colorclass 2.2.0 (now a package)

__version__ = '0.04'

#------------------------------------------------------------------------------
# TODO:
# - several styles
# - colorized rows or cells
# - automatic width for the last column, based on max total width
# - automatic width for selected columns, based on N first lines
# - determine the console width

# === IMPORTS =================================================================

import textwrap
import sys, os

from thirdparty import colorclass

# On Windows, colorclass needs to be enabled:
if os.name == 'nt':
    colorclass.Windows.enable(auto_colors=True)


# === CLASSES =================================================================


class TableStyle(object):
    """
    Style for a TableStream.
    This base class can be derived to create new styles.
    Default style:
    +------+---+
    |Header|   +
    +------+---+
    |      |   |
    +------+---+
    """
    # Header rows:
    header_top = True
    header_top_left = '+'
    header_top_horiz = '-'
    header_top_middle = '+'
    header_top_right = '+'

    header_vertical_left = '|'
    header_vertical_middle = '|'
    header_vertical_right = '|'

    # Separator line between header and normal rows:
    header_sep = True
    header_sep_left = '+'
    header_sep_horiz = '-'
    header_sep_middle = '+'
    header_sep_right = '+'

    # Top row if there is no header:
    noheader_top = True
    noheader_top_left = '+'
    noheader_top_horiz = '-'
    noheader_top_middle = '+'
    noheader_top_right = '+'

    # Normal rows
    vertical_left = '|'
    vertical_middle = '|'
    vertical_right = '|'

    # Separator line between rows:
    sep = False
    sep_left = '+'
    sep_horiz = '-'
    sep_middle = '+'
    sep_right = '+'

    # Bottom line
    bottom = True
    bottom_left = '+'
    bottom_horiz = '-'
    bottom_middle = '+'
    bottom_right = '+'


class TableStyleSlim(object):
    """
    Style for a TableStream.
    Example:
    ------+---
    Header|
    ------+---
          |
    ------+---
    """
    # Header rows:
    header_top = True
    header_top_left = ''
    header_top_horiz = '-'
    header_top_middle = '+'
    header_top_right = ''

    header_vertical_left = ''
    header_vertical_middle = '|'
    header_vertical_right = ''

    # Separator line between header and normal rows:
    header_sep = True
    header_sep_left = ''
    header_sep_horiz = '-'
    header_sep_middle = '+'
    header_sep_right = ''

    # Top row if there is no header:
    noheader_top = True
    noheader_top_left = ''
    noheader_top_horiz = '-'
    noheader_top_middle = '+'
    noheader_top_right = ''

    # Normal rows
    vertical_left = ''
    vertical_middle = '|'
    vertical_right = ''

    # Separator line between rows:
    sep = False
    sep_left = ''
    sep_horiz = '-'
    sep_middle = '+'
    sep_right = ''

    # Bottom line
    bottom = True
    bottom_left = ''
    bottom_horiz = '-'
    bottom_middle = '+'
    bottom_right = ''



class TableStream(object):
    """
    a TableStream object can format table data for pretty printing as text,
    to be displayed on the console or written to any file-like object.
    The table data can be provided as rows, each row is an iterable of
    cells. The text in each cell is wrapped to fit into a maximum width
    set for each column.
    Contrary to many table pretty printing libraries, TableStream writes
    each row to the output as soon as it is provided, and the whole table
    does not need to be built in memory before printing.
    It is therefore suitable for large tables, or tables that take time to
    be processed row by row.
    """

    def __init__(self, column_width, header_row=None, style=TableStyle, outfile=sys.stdout):
        self.column_width = column_width
        self.num_columns = len(column_width)
        self.header_row = header_row
        assert (header_row is None) or len(header_row) == self.num_columns
        self.style = style
        self.outfile = outfile
        if header_row is not None:
            self.write_header()
        elif self.style.noheader_top:
            self.write_noheader_top()


    def write(self, s):
        """
        shortcut for self.outfile.write()
        """
        self.outfile.write(s)

    def write_row(self, row, last=False, colors=None):
        assert len(row) == self.num_columns
        columns = []
        max_lines = 0
        for i in xrange(self.num_columns):
            cell = row[i]
            # Convert to string:
            # TODO: handle unicode properly
            # TODO: use only unicode for textwrapper, to avoid str length issues
            if isinstance(cell, bytes):
                # encode to UTF8, avoiding errors
                cell = cell.decode('utf-8', errors='replace')
            else:
                cell = unicode(cell)
            # Wrap cell text according to the column width
            # TODO: use a TextWrapper object for each column instead
            column = textwrap.wrap(cell, width=self.column_width[i])
            # apply colors to each line of the cell if needed:
            if colors is not None and self.outfile.isatty():
                color = colors[i]
                if color:
                    for j in xrange(len(column)):
                        # print '%r: %s' % (column[j], type(column[j]))
                        column[j] = colorclass.Color('{auto%s}%s{/%s}' % (color, column[j], color))
            columns.append(column)
            # determine which column has the highest number of lines
            max_lines = max(len(columns[i]), max_lines)
        # transpose: write output line by line
        for j in xrange(max_lines):
            self.write(self.style.vertical_left)
            for i in xrange(self.num_columns):
                column = columns[i]
                if j<len(column):
                    # text to be written
                    text_width = len(column[j])
                    self.write(column[j] + ' '*(self.column_width[i]-text_width))
                else:
                    # no more lines for this column
                    # TODO: precompute empty cells once
                    self.write(' '*(self.column_width[i]))
                if i < (self.num_columns - 1):
                    self.write(self.style.vertical_middle)
            self.write(self.style.vertical_right)
            self.write('\n')
        if self.style.sep and not last:
            self.write_sep()

    def make_line(self, left, horiz, middle, right):
        """
        build a line based on the provided elements
        example: '+---+--+-------+'
        :param left:
        :param horiz:
        :param middle:
        :param right:
        :return:
        """
        return left + middle.join([horiz * width for width in self.column_width]) + right + '\n'

    def write_header_top(self):
        s = self.style
        line = self.make_line(left=s.header_top_left, horiz=s.header_top_horiz,
                              middle=s.header_top_middle, right=s.header_top_right)
        self.write(line)

    def write_header_sep(self):
        s = self.style
        line = self.make_line(left=s.header_sep_left, horiz=s.header_sep_horiz,
                              middle=s.header_sep_middle, right=s.header_sep_right)
        self.write(line)

    def write_header(self):
        if self.style.header_top:
            self.write_header_top()
        self.write_row(self.header_row)
        if self.style.header_sep:
            self.write_header_sep()

    def write_noheader_top(self):
        s = self.style
        line = self.make_line(left=s.noheader_top_left, horiz=s.noheader_top_horiz,
                              middle=s.noheader_top_middle, right=s.noheader_top_right)
        self.write(line)

    def write_sep(self):
        s = self.style
        line = self.make_line(left=s.sep_left, horiz=s.sep_horiz,
                              middle=s.sep_middle, right=s.sep_right)
        self.write(line)

    def write_bottom(self):
        s = self.style
        line = self.make_line(left=s.bottom_left, horiz=s.bottom_horiz,
                              middle=s.bottom_middle, right=s.bottom_right)
        self.write(line)

    def close(self):
        self.write_bottom()


if __name__ == '__main__':
    t = TableStream([10, 5, 20], header_row=['i', 'i*i', '2**i'], style=TableStyleSlim)
    t.write_row(['test', 'test', 'test'])
    cell = 'a very very long text'
    t.write_row([cell, cell, cell], colors=['blue', None, 'red'])
    for i in range(1, 11):
        t.write_row([i, i*i, 2**i])
    t.close()


