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
# 2017-10-23       PL: - add check for fldSimple codes
# 2017-10-24       PL: - group tags and track begin/end tags to keep DDE strings together

__version__ = '0.52dev2'

#------------------------------------------------------------------------------
# TODO: detect beginning/end of fields, to separate each field
# TODO: test if DDE links can also appear in headers, footers and other places
# TODO: field codes can be in headers/footers/comments - parse these
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


# === CONSTANTS ==============================================================


NS_WORD = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NO_QUOTES = False
# XML tag for 'w:instrText'
TAG_W_INSTRTEXT = '{%s}instrText' % NS_WORD
TAG_W_FLDSIMPLE = '{%s}fldSimple' % NS_WORD
TAG_W_FLDCHAR = '{%s}fldChar' % NS_WORD
TAG_W_P = "{%s}p" % NS_WORD
TAG_W_R = "{%s}r" % NS_WORD
ATTR_W_INSTR = '{%s}instr' % NS_WORD
ATTR_W_FLDCHARTYPE = '{%s}fldCharType' % NS_WORD

LOCATIONS = ['word/document.xml','word/endnotes.xml','word/footnotes.xml','word/header1.xml','word/footer1.xml','word/header2.xml','word/footer2.xml','word/comments.xml']
# === FUNCTIONS ==============================================================

def process_args():
    parser = argparse.ArgumentParser(description='A python tool to detect and extract DDE links in MS Office files')
    parser.add_argument("filepath", help="path of the file to be analyzed")
    parser.add_argument("--nounquote", help="don't unquote values",action='store_true')
    args = parser.parse_args()

    if not os.path.exists(args.filepath):
        print('File {} does not exist.'.format(args.filepath))
        sys.exit(1)

    return args



def process_file(data):
   
    # parse the XML data:
    root = ET.fromstring(data)
    fields = []
    ddetext = u''
    level = 0
    # find all the tags 'w:p':
    # parse each for begin and end tags, to group DDE strings
    # fldChar can be in either a w:r element, floating alone in the w:p or spread accross w:p tags
    # escape DDE if quoted etc
    # (each is a chunk of a DDE link)
    for subs in root.iter(TAG_W_P):
        elem = None
        for e in subs:
            #check if w:r and if it is parse children elements to pull out the first FLDCHAR or INSTRTEXT
            if e.tag == TAG_W_R:
                for child in e:
                    if child.tag == TAG_W_FLDCHAR or child.tag == TAG_W_INSTRTEXT:
                        elem = child
                        break
            else:
                elem = e
            #this should be an error condition
            if elem is None:
                continue
    
            #check if FLDCHARTYPE and whether "begin" or "end" tag
            if elem.attrib.get(ATTR_W_FLDCHARTYPE) is not None:
                if elem.attrib[ATTR_W_FLDCHARTYPE] == "begin":
                    level += 1    
                if elem.attrib[ATTR_W_FLDCHARTYPE] == "end":
                    level -= 1
                    if level == 0 or level == -1 : # edge-case where level becomes -1
                        fields.append(ddetext)
                        ddetext = u''
                        level = 0 # reset edge-case
        
            # concatenate the text of the field, if present:
            if elem.tag == TAG_W_INSTRTEXT and elem.text is not None:
                #expand field code if QUOTED
                ddetext += unquote(elem.text)


    for elem in root.iter(TAG_W_FLDSIMPLE):
        # concatenate the attribute of the field, if present:
        if elem.attrib is not None:
            fields.append(elem.attrib[ATTR_W_INSTR])
    

    return fields

def unquote(field): 
    if "QUOTE" not in field or NO_QUOTES:
        return field
    #split into components
    parts = field.strip().split(" ")
    ddestr = ""
    for p in parts[1:]:
        try: 
             ch = chr(int(p))
        except ValueError:
            ch = p
        ddestr += ch 
    return ddestr

#=== MAIN =================================================================

def main():
    # print banner with version
    print ('msodde %s - http://decalage.info/python/oletools' % __version__)
    print ('THIS IS WORK IN PROGRESS - Check updates regularly!')
    print ('Please report any issue at https://github.com/decalage2/oletools/issues')
    print ('')

    args = process_args()
    print('Opening file: %s' % args.filepath)
    if args.nounquote :
        global NO_QUOTES
        NO_QUOTES = True
    z = zipfile.ZipFile(args.filepath)
    for filepath in z.namelist():
        if filepath in LOCATIONS:
            data = z.read(filepath)
            fields = process_file(data)
            if len(fields) > 0:
                print ('DDE Links in %s:'%filepath)
                for f in fields:
                    print(f)
    z.close()
    


if __name__ == '__main__':
    main()
