""" Basic tests for ooxml.py """

import unittest

import os
from os.path import join, splitext
from tests.test_utils import DATA_BASE_DIR
from oletools.thirdparty.olefile import isOleFile
from oletools import ooxml


class TestOOXML(unittest.TestCase):
    """ Tests my cool new feature """

    def test_all_rough(self):
        """Checks all samples, expect either ole files or good ooxml output"""
        acceptable = ooxml.DOCTYPE_EXCEL, ooxml.DOCTYPE_WORD, \
                     ooxml.DOCTYPE_POWERPOINT
        except_files = 'empty', 'text'
        except_extns = '.xml', '.rtf'
        for base_dir, _, files in os.walk(DATA_BASE_DIR):
            for filename in files:
                if filename in except_files:
                    #print('skip file: ' + filename)
                    continue
                if splitext(filename)[1] in except_extns:
                    #print('skip extn: ' + filename)
                    continue

                full_name = join(base_dir, filename)
                if isOleFile(full_name):
                    #print('skip ole: ' + filename)
                    continue
                try:
                    doctype = ooxml.get_type(full_name)
                except Exception:
                    self.fail('Failed to get doctype of {0}'.format(filename))
                self.assertIn(doctype, acceptable)
                #print('ok: ' + filename + doctype)


# just in case somebody calls this file as a script
if __name__ == '__main__':
    unittest.main()
