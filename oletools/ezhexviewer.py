#!/usr/bin/env python
"""
ezhexviewer.py

A simple hexadecimal viewer based on easygui. It should work on any platform
with Python 2.x or 3.x.

Usage: ezhexviewer.py [file]

Usage in a python application:

    import ezhexviewer
    ezhexviewer.hexview_file(filename)
    ezhexviewer.hexview_data(data)


ezhexviewer project website: http://www.decalage.info/python/ezhexviewer

ezhexviewer is copyright (c) 2012-2019, Philippe Lagadec (http://www.decalage.info)
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

#------------------------------------------------------------------------------
# CHANGELOG:
# 2012-09-17 v0.01 PL: - first version
# 2012-10-04 v0.02 PL: - added license
# 2016-09-06 v0.50 PL: - added main function for entry points in setup.py
# 2016-10-26       PL: - fixed to run on Python 2+3
# 2017-03-23 v0.51 PL: - fixed display of control characters (issue #151)
# 2017-04-26       PL: - fixed absolute imports (issue #141)
# 2018-09-15 v0.54 PL: - easygui is now a dependency

__version__ = '0.54'

#-----------------------------------------------------------------------------
# TODO:
# + options to set title and msg

# === IMPORTS ================================================================

import sys, os

# IMPORTANT: it should be possible to run oletools directly as scripts
# in any directory without installing them with pip or setup.py.
# In that case, relative imports are NOT usable.
# And to enable Python 2+3 compatibility, we need to use absolute imports,
# so we add the oletools parent folder to sys.path (absolute+normalized path):
_thismodule_dir = os.path.normpath(os.path.abspath(os.path.dirname(__file__)))
# print('_thismodule_dir = %r' % _thismodule_dir)
_parent_dir = os.path.normpath(os.path.join(_thismodule_dir, '..'))
# print('_parent_dir = %r' % _thirdparty_dir)
if not _parent_dir in sys.path:
    sys.path.insert(0, _parent_dir)

import easygui

# === PYTHON 2+3 SUPPORT ======================================================

if sys.version_info[0] >= 3:
    # Python 3 specific adaptations
    # py3 range = py2 xrange
    xrange = range
    PYTHON3 = True
else:
    PYTHON3 = False

def xord(char):
    '''
    workaround for ord() to work on characters from a bytes string with
    Python 2 and 3. If s is a bytes string, s[i] is a bytes string of
    length 1 on Python 2, but it is an integer on Python 3...
    xord(c) returns ord(c) if c is a bytes string, or c if it is already
    an integer.
    :param char: int or bytes of length 1
    :return: ord(c) if bytes, c if int
    '''
    if isinstance(char, int):
        return char
    else:
        return ord(char)

def bchr(x):
    '''
    workaround for chr() to return a bytes string of length 1 with
    Python 2 and 3. On Python 3, chr returns a unicode string, but
    on Python 2 it is a bytes string.
    bchr() always returns a bytes string on Python 2+3.
    :param x: int
    :return: chr(x) as a bytes string
    '''
    if PYTHON3:
        # According to the Python 3 documentation, bytes() can be
        # initialized with an iterable:
        return bytes([x])
    else:
        return chr(x)

#------------------------------------------------------------------------------
# The following code (hexdump3 only) is a modified version of the hex dumper
# recipe published on ASPN by Sebastien Keim and Raymond Hattinger under the
# PSF license. I added the startindex parameter.
# see http://aspn.activestate.com/ASPN/Cookbook/Python/Recipe/142812
# PSF license: http://docs.python.org/license.html
# Copyright (c) 2001-2012 Python Software Foundation; All Rights Reserved

FILTER = b''.join([(len(repr(bchr(x)))<=4 and x>=0x20) and bchr(x) or b'.' for x in range(256)])

def hexdump3(src, length=8, startindex=0):
    """
    Returns a hexadecimal dump of a binary string.
    length: number of bytes per row.
    startindex: index of 1st byte.
    """
    result=[]
    # pylint: disable-next=possibly-used-before-assignment
    for i in xrange(0, len(src), length):
        s = src[i:i+length]
        hexa = ' '.join(["%02X" % xord(x) for x in s])
        printable = s.translate(FILTER)
        if PYTHON3:
            # On Python 3, need to convert printable from bytes to str:
            printable = printable.decode('latin1')
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


# === MAIN ===================================================================

def main():
    try:
        filename = sys.argv[1]
    except:
        filename = easygui.fileopenbox()
    if filename:
        try:
            hexview_file(filename, msg='File: %s' % filename)
        except:
            easygui.exceptionbox(msg='Error:', title='ezhexviewer')


if __name__ == '__main__':
    main()
