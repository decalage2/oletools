#!/usr/bin/env python
"""
msodde.py

msodde is a script to parse MS Office documents
(e.g. Word, Excel), to detect and extract DDE links.

Supported formats:
- Word 2007+ (.docx, .dotx, .docm, .dotm)

Author: Philippe Lagadec - http://www.decalage.info
License: BSD, see source code or documentation

msodde is part of the python-oletools package:
http://www.decalage.info/python/oletools
"""

# === LICENSE ==================================================================

# msodde is copyright (c) 2017 Philippe Lagadec (http://www.decalage.info)
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

from __future__ import print_function

#------------------------------------------------------------------------------
# CHANGELOG:
# 2017-10-18 v0.52 PL: - first version
# 2017-10-20       PL: - fixed issue #202 (handling empty xml tags)
# 2017-10-25       CH: - add json output

__version__ = '0.52dev2'

#------------------------------------------------------------------------------
# TODO: detect beginning/end of fields, to separate each field
# TODO: test if DDE links can also appear in headers, footers and other places
# TODO: add xlsx support

#------------------------------------------------------------------------------
# REFERENCES:


#--- IMPORTS ------------------------------------------------------------------

# import lxml or ElementTree for XML parsing:
try:
    # lxml: best performance for XML processing
    import lxml.etree as ET
except ImportError:
    import xml.etree.cElementTree as ET

import argparse
import zipfile
import os
import sys
import json


# === CONSTANTS ==============================================================


NS_WORD = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

# XML tag for 'w:instrText'
TAG_W_INSTRTEXT = '{%s}instrText' % NS_WORD
TAG_W_FLDSIMPLE = '{%s}fldSimple' % NS_WORD
TAG_W_INSTRATTR= '{%s}instr' % NS_WORD

# banner to be printed at program start
BANNER = """
msodde %s - http://decalage.info/python/oletools
THIS IS WORK IN PROGRESS - Check updates regularly!
Please report any issue at https://github.com/decalage2/oletools/issues
""" % __version__

BANNER_JSON = dict(type='meta', version=__version__, name='msodde',
                   link='http://decalage.info/python/oletools',
                   message='THIS IS WORK IN PROGRESS - Check updates regularly! '
                            'Please report any issue at '
                            'https://github.com/decalage2/oletools/issues')

# === FUNCTIONS ==============================================================

def process_args():
    parser = argparse.ArgumentParser(description='A python tool to detect and extract DDE links in MS Office files')
    parser.add_argument("filepath", help="path of the file to be analyzed")
    parser.add_argument("--json", '-j', action='store_true',
                        help="Output in json format")

    args = parser.parse_args()

    if not os.path.exists(args.filepath):
        print('File {} does not exist.'.format(args.filepath))
        sys.exit(1)

    return args



def process_file(filepath):
    z = zipfile.ZipFile(filepath)
    data = z.read('word/document.xml')
    z.close()
    # parse the XML data:
    root = ET.fromstring(data)
    text = u''
    # find all the tags 'w:instrText':
    # (each is a chunk of a DDE link)
    for elem in root.iter(TAG_W_INSTRTEXT):
        # concatenate the text of the field, if present:
        if elem.text is not None:
            text += elem.text

    for elem in root.iter(TAG_W_FLDSIMPLE):
        # concatenate the attribute of the field, if present:
        if elem.attrib is not None:
            text += elem.attrib[TAG_W_INSTRATTR]
    

    return text


#=== MAIN =================================================================

def main():
    args = process_args()

    if args.json:
        jout = []
        jout.append(BANNER_JSON)
    else:
        # print banner with version
        print(BANNER)

    if not args.json:
        print('Opening file: %s' % args.filepath)

    text = ''
    return_code = 1
    try:
        text = process_file(args.filepath)
        return_code = 0
    except Exception as exc:
        if args.json:
            jout.append(dict(type='error', error=type(exc).__name__,
                             message=str(exc)))  # strange: str(exc) is enclosed in ""
        else:
            raise

    if args.json:
        for line in text.splitlines():
            jout.append(dict(type='dde-link', link=line.strip()))
        json.dump(jout, sys.stdout, check_circular=False, indent=4)
        print()   # add a newline after closing "]"
        sys.exit(return_code)  # required if we catch an exception in json-mode
    else:
        print ('DDE Links:')
        print(text)


if __name__ == '__main__':
    main()
