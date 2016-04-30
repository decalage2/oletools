#!/usr/bin/env python
"""
ezhexviewer.py

A simple hexadecimal viewer based on easygui. It should work on any platform
with Python 2.x.

Usage: ezhexviewer.py [file]

Usage in a python application:

    import ezhexviewer
    ezhexviewer.hexview_file(filename)
    ezhexviewer.hexview_data(data)


ezhexviewer project website: http://www.decalage.info/python/ezhexviewer

ezhexviewer is copyright (c) 2012-2015, Philippe Lagadec (http://www.decalage.info)
All rights reserved.

Redistribution and use in source and binary forms, with or without modification,
are permitted provided that the following conditions are met:

 * Redistributions of source code must retain the above copyright notice, this
   list of conditions and the following disclaimer.
 * Redistributions in binary form must reproduce the above copyright notice,
   this list of conditions and the following disclaimer in the documentation
   and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
"""

__version__ = '0.02'

#------------------------------------------------------------------------------
# CHANGELOG:
# 2012-09-17 v0.01 PL: - first version
# 2012-10-04 v0.02 PL: - added license

#------------------------------------------------------------------------------
# TODO:
# + options to set title and msg


from thirdparty.easygui import easygui
import sys

#------------------------------------------------------------------------------
# The following code (hexdump3 only) is a modified version of the hex dumper
# recipe published on ASPN by Sebastien Keim and Raymond Hattinger under the
# PSF license. I added the startindex parameter.
# see http://aspn.activestate.com/ASPN/Cookbook/Python/Recipe/142812
# PSF license: http://docs.python.org/license.html
# Copyright (c) 2001-2012 Python Software Foundation; All Rights Reserved

FILTER=''.join([(len(repr(chr(x)))==3) and chr(x) or '.' for x in range(256)])

def hexdump3(src, length=8, startindex=0):
    """
    Returns a hexadecimal dump of a binary string.
    length: number of bytes per row.
    startindex: index of 1st byte.
    """
    result=[]
    for i in xrange(0, len(src), length):
       s = src[i:i+length]
       hexa = ' '.join(["%02X"%ord(x) for x in s])
       printable = s.translate(FILTER)
       result.append("%08X   %-*s   %s\n" % (i+startindex, length*3, hexa, printable))
    return ''.join(result)

# end of PSF-licensed code.
#------------------------------------------------------------------------------


def hexview_data (data, msg='', title='ezhexviewer', length=16, startindex=0):
    hex = hexdump3(data, length=length, startindex=startindex)
    easygui.codebox(msg=msg, title=title, text=hex)


def hexview_file (filename, msg='', title='ezhexviewer', length=16, startindex=0):
    data = open(filename, 'rb').read()
    hexview_data(data, msg=msg, title=title, length=length, startindex=startindex)


if __name__ == '__main__':
    try:
        filename = sys.argv[1]
    except:
        filename = easygui.fileopenbox()
    if filename:
        try:
            hexview_file(filename, msg='File: %s' % filename)
        except:
            easygui.exceptionbox(msg='Error:', title='ezhexviewer')
