#!/usr/bin/env python
"""
exportvba.py

Exports VBA from files supplied on the commane line. Output files will be
placed in folders next to each input file.
"""

import os, shutil, sys
from olevba3 import VBA_Parser, TYPE_OLE, TYPE_OpenXML, TYPE_Word2003_XML, TYPE_MHTML, filter_vba

def export(inFile):
    vbaparser = VBA_Parser(inFile)
    if not vbaparser.detect_vba_macros():
        return

    inFile = os.path.relpath(inFile)
    outDir = os.path.splitext(inFile)[0] + os.sep

    if os.path.exists(outDir):
        try:
            shutil.rmtree(outDir)
        except OSError:
            raise
    try:
        os.makedirs(outDir)
    except OSError:
        raise

    print(inFile + ':')
    for (filename, stream_path, vba_filename, vba_code) in vbaparser.extract_macros():
        if isinstance(vba_code, bytes):
            vba_code = vba_code.decode('utf-8', 'backslashreplace')
        vba_code = filter_vba(vba_code)
        if not vba_code.strip():
            continue

        outFile = outDir + vba_filename + '.vba'
        with open(outFile, 'w') as f:
            f.write(vba_code)

        print('    ' + outFile)

def main(argv):
    for x in argv[1:]:
        print('')
        export(x)

if __name__ == '__main__':
    main(sys.argv)
