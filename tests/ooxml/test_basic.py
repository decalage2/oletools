""" Basic tests for ooxml.py """

import unittest

import os
from os.path import join, splitext
from tests.test_utils import DATA_BASE_DIR
from oletools.thirdparty.olefile import isOleFile
from oletools import ooxml


class TestOOXML(unittest.TestCase):
    """ Tests correct detection of doc type """

    DO_DEBUG = False

    def test_all_rough(self):
        """Checks all samples, expect either ole files or good ooxml output"""
        # map from extension to expected doctype
        ext2doc = dict(
            docx=ooxml.DOCTYPE_WORD, docm=ooxml.DOCTYPE_WORD,
            xml=(ooxml.DOCTYPE_EXCEL_XML, ooxml.DOCTYPE_WORD_XML),
            xlsx=ooxml.DOCTYPE_EXCEL, xlsm=ooxml.DOCTYPE_EXCEL,
            xlsb=ooxml.DOCTYPE_EXCEL,
            pptx=ooxml.DOCTYPE_POWERPOINT, pptm=ooxml.DOCTYPE_POWERPOINT,
        )

        # files that are neither OLE nor xml:
        except_files = 'empty', 'text'
        except_extns = 'rtf'

        # analyse all files in data dir
        for base_dir, _, files in os.walk(DATA_BASE_DIR):
            for filename in files:
                if filename in except_files:
                    if self.DO_DEBUG:
                        print('skip file: ' + filename)
                    continue
                extn = splitext(filename)[1]
                if extn:
                    extn = extn[1:]      # remove the dot
                if extn in except_extns:
                    if self.DO_DEBUG:
                        print('skip extn: ' + filename)
                    continue

                full_name = join(base_dir, filename)
                if isOleFile(full_name):
                    if self.DO_DEBUG:
                        print('skip ole: ' + filename)
                    continue
                acceptable = ext2doc[extn]
                if not isinstance(acceptable, tuple):
                    acceptable = (acceptable, )
                try:
                    doctype = ooxml.get_type(full_name)
                except Exception:
                    self.fail('Failed to get doctype of {0}'.format(filename))
                self.assertTrue(doctype in acceptable,
                                msg='Doctype "{0}" for {1} not acceptable'
                                    .format(doctype, full_name))
                if self.DO_DEBUG:
                    print('ok: {0} --> {1}'.format(filename, doctype))


# just in case somebody calls this file as a script
if __name__ == '__main__':
    unittest.main()
