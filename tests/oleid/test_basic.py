"""
Test basic functionality of oleid

Should work with python2 and python3!
"""

import unittest
import os
from os.path import join, relpath, splitext
from oletools import oleid
from oletools.ftguess import CONTAINER

from tests.test_utils.testdata_reader import loop_over_files, DATA_BASE_DIR


class TestOleIDBasic(unittest.TestCase):
    """Test basic functionality of OleID"""

    def setUp(self):
        """Called before tests; populates self.oleids"""
        self.oleids = []
        for filename, file_contents in loop_over_files():
            curr_id = oleid.OleID(filename=filename, data=file_contents)
            value_dict = dict((ind.id, ind.value) for ind in curr_id.check())
            self.oleids.append((filename, value_dict))

    # note: indicators "ftype" and "container" are from ftguess,
    #       so tested there, already

    def test_properties(self):
        """Test indicators "appname", "codepage" and "author" of ole files."""
        for filename, value_dict in self.oleids:
            # print('Debugging: testing file {0}'.format(filename))
            if value_dict['container'] != CONTAINER.OLE:
                self.assertNotIn('appname', value_dict)
                self.assertNotIn('codepage', value_dict)
                self.assertNotIn('author', value_dict)
                continue

            before_dot, suffix = splitext(filename)
            if suffix == '.zip':
                suffix = splitext(before_dot)[1]

            if 'encrypted' in filename \
                    and suffix != '.xls' and suffix != '.doc':
                self.assertEqual(value_dict['appname'], None)
                self.assertEqual(value_dict['codepage'], None)
                self.assertEqual(value_dict['author'], None)
                continue

            if suffix.startswith('.d'):
                self.assertEqual(value_dict['appname'],
                                 b'Microsoft Office Word')
            elif suffix.startswith('.x'):
                self.assertIn(value_dict['appname'],
                              (b'Microsoft Office Excel', b'Microsoft Excel'))
                # old types have no "Office" in the app name
            elif suffix.startswith('.p'):
                self.assertEqual(value_dict['appname'],
                                 b'Microsoft Office PowerPoint')
            else:
                self.fail('Unexpected suffix {0} from app {1}'
                          .format(suffix, value_dict['appname']))

            if 'utf_16le-korean' in filename:
                self.assertEqual(value_dict['codepage'],
                                 '949: ANSI/OEM Korean (Unified Hangul Code)')
                self.assertEqual(value_dict['author'],
                                 b'\xb1\xe8\xb1\xe2\xc1\xa4;kijeong')
            else:
                self.assertEqual(value_dict['codepage'],
                                 '1252: ANSI Latin 1; Western European (Windows)')
                self.assertIn(value_dict['author'],
                              (b'user', b'schulung',
                               b'xxxxxxxxxxxx', b'zzzzzzzzzzzz'))

    def test_encrypted(self):
        """Test indicator "encrypted"."""
        for filename, value_dict in self.oleids:
            # print('Debugging: testing file {0}'.format(filename))
            self.assertEqual(value_dict['encrypted'], 'encrypted' in filename)

    def test_external_rels(self):
        """Test indicator for external relationships."""
        for filename, value_dict in self.oleids:
            # print('Debugging: testing file {0}'.format(filename))
            self.assertEqual(value_dict['ext_rels'],
                             '/external_link/' in filename)

    def test_objectpool(self):
        """Test indicator for ObjectPool stream in ole files."""
        for filename, value_dict in self.oleids:
            # print('Debugging: testing file {0}'.format(filename))
            if (filename.startswith('oleobj/sample_with_')
                        or filename.startswith('oleobj/embedded')) \
                    and (filename.endswith('.doc') 
                         or filename.endswith('.dot')):
                self.assertTrue(value_dict['ObjectPool'])
            else:
                self.assertFalse(value_dict['ObjectPool'])

    def test_macros(self):
        """Test indicator for macros."""
        for filename, value_dict in self.oleids:
            # TODO: we need a sample file with xlm macros
            before_dot, suffix = splitext(filename)
            if suffix == '.zip':
                suffix = splitext(before_dot)[1]
            # print('Debugging: {1}, {2} for {0}'
            #       .format(filename, value_dict['vba'], value_dict['xlm']))

            # xlm detection does not work in-memory (yet)
            # --> xlm is "unknown" for excel files, except some encrypted files
            self.assertIn(value_dict['xlm'], ('Unknown', 'No'))

            # "macro detection" in text files leads to interesting results:
            if filename in ('ooxml/dde-in-excel2003.xml',    # not really
                            'encrypted/autostart-encrypt-standardpassword.xls',
                            'msodde/dde-in-csv.csv',     # "Windows" "calc.exe"
                            'msodde/dde-in-excel2003.xml',   # same as above
                            'oleform/oleform-PR314.docm',
                            'basic/empty',                   # WTF?
                            'basic/text'):                   # no macros!
                self.assertEqual(value_dict['vba'], 'Yes')
            else:
                self.assertEqual(value_dict['vba'], 'No')

    def test_flash(self):
        """Test indicator for flash."""
        # TODO: add a sample that contains flash
        for filename, value_dict in self.oleids:
            # print('Debugging: testing file {0}'.format(filename))
            self.assertEqual(value_dict['flash'], 0)



# just in case somebody calls this file as a script
if __name__ == '__main__':
    unittest.main()
