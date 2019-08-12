#!/usr/bin/env python
"""
olebrowse.py

A simple GUI to browse OLE files (e.g. MS Word, Excel, Powerpoint documents), to
view and extract individual data streams.

Usage: olebrowse.py [file]

olebrowse project website: http://www.decalage.info/python/olebrowse

olebrowse is part of the python-oletools package:
http://www.decalage.info/python/oletools

olebrowse is copyright (c) 2012-2019, Philippe Lagadec (http://www.decalage.info)
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
# 2014-11-29 v0.02 PL: - use olefile instead of OleFileIO_PL
# 2017-04-26 v0.51 PL: - fixed absolute imports (issue #141)
# 2018-09-11 v0.54 PL: - olefile is now a dependency

__version__ = '0.54'

#------------------------------------------------------------------------------
# TODO:
# - menu option to open another file
# - menu option to display properties
# - menu option to run other oletools, external tools such as OfficeCat?
# - for a stream, display info: size, path, etc
# - stream info: magic, entropy, ... ?

# === IMPORTS ================================================================

import optparse, sys, os

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
import olefile
from oletools import ezhexviewer

ABOUT = '~ About olebrowse'
QUIT  = '~ Quit'


def about ():
    """
    Display information about this tool
    """
    easygui.textbox(title='About olebrowse', text=__doc__)


def browse_stream (ole, stream):
    """
    Browse a stream (hex view or save to file)
    """
    #print 'stream:', stream
    while True:
        msg ='Select an action for the stream "%s", or press Esc to exit' % repr(stream)
        actions = [
            'Hex view',
##                'Text view',
##                'Repr view',
            'Save stream to file',
            '~ Back to main menu',
            ]
        action = easygui.choicebox(msg, title='olebrowse', choices=actions)
        if action is None or 'Back' in action:
            break
        elif action.startswith('Hex'):
            data = ole.openstream(stream).getvalue()
            ezhexviewer.hexview_data(data, msg='Stream: %s' % stream, title='olebrowse')
##            elif action.startswith('Text'):
##                data = ole.openstream(stream).getvalue()
##                easygui.codebox(title='Text view - %s' % stream, text=data)
##            elif action.startswith('Repr'):
##                data = ole.openstream(stream).getvalue()
##                easygui.codebox(title='Repr view - %s' % stream, text=repr(data))
        elif action.startswith('Save'):
            data = ole.openstream(stream).getvalue()
            fname = easygui.filesavebox(default='stream.bin')
            if fname is not None:
                f = open(fname, 'wb')
                f.write(data)
                f.close()
                easygui.msgbox('stream saved to file %s' % fname)



def main():
    """
    Main function
    """
    try:
        filename = sys.argv[1]
    except:
        filename = easygui.fileopenbox()
    try:
        ole = olefile.OleFileIO(filename)
        listdir = ole.listdir()
        streams = []
        for direntry in listdir:
            #print direntry
            streams.append('/'.join(direntry))
        streams.append(ABOUT)
        streams.append(QUIT)
        stream = True
        while stream is not None:
            msg ="Select a stream, or press Esc to exit"
            title = "olebrowse"
            stream = easygui.choicebox(msg, title, streams)
            if stream is None or stream == QUIT:
                break
            if stream == ABOUT:
                about()
            else:
                browse_stream(ole, stream)
    except:
        easygui.exceptionbox()




if __name__ == '__main__':
    main()
