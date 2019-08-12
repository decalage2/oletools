"""
Test basic functionality of oleid

Should work with python2 and python3!
"""

import unittest
import os
from os.path import join, relpath, splitext
from oletools import oleid

# Directory with test data, independent of current working directory
from tests.test_utils import DATA_BASE_DIR


class TestOleIDBasic(unittest.TestCase):
    """Test basic functionality of OleID"""

    def test_all(self):
        """Run all file in test-data through oleid and compare to known ouput"""
        # this relies on order of indicators being constant, could relax that
        # Also requires that files have the correct suffixes (no rtf in doc)
        NON_OLE_SUFFIXES = ('.xml', '.csv', '.rtf', '', '.odt', '.ods', '.odp')
        NON_OLE_VALUES = (False, )
        WORD = b'Microsoft Office Word'
        PPT = b'Microsoft Office PowerPoint'
        EXCEL = b'Microsoft Excel'
        CRYPT = (True, False, 'unknown', True, False, False, False, False,
                 False, False, 0)
        OLE_VALUES = {
            'oleobj/sample_with_lnk_file.doc': (True, True, WORD, False, True,
                                                False, False, False, False,
                                                True, 0),
            'oleobj/embedded-simple-2007.xlsb': (False,),
            'oleobj/embedded-simple-2007.docm': (False,),
            'oleobj/embedded-simple-2007.xltx': (False,),
            'oleobj/embedded-simple-2007.xlam': (False,),
            'oleobj/embedded-simple-2007.dotm': (False,),
            'oleobj/sample_with_lnk_file.ppt': (True, True, PPT, False, False,
                                                False, False, True, False,
                                                False, 0),
            'oleobj/embedded-simple-2007.xlsx': (False,),
            'oleobj/embedded-simple-2007.xlsm': (False,),
            'oleobj/embedded-simple-2007.ppsx': (False,),
            'oleobj/embedded-simple-2007.pps': (True, True, PPT, False, False,
                                                False, False, True, False,
                                                False, 0),
            'oleobj/embedded-simple-2007.xla': (True, True, EXCEL, False,
                                                False, False, True, False,
                                                False, False, 0),
            'oleobj/sample_with_calc_embedded.doc': (True, True, WORD, False,
                                                     True, False, False, False,
                                                     False, True, 0),
            'oleobj/embedded-unicode-2007.docx': (False,),
            'oleobj/embedded-unicode.doc': (True, True, WORD, False, True,
                                            False, False, False, False, True,
                                            0),
            'oleobj/embedded-simple-2007.doc': (True, True, WORD, False, True,
                                                False, False, False, False,
                                                True, 0),
            'oleobj/embedded-simple-2007.xls': (True, True, EXCEL, False,
                                                False, False, True, False,
                                                False, False, 0),
            'oleobj/embedded-simple-2007.dot': (True, True, WORD, False, True,
                                                False, False, False, False,
                                                True, 0),
            'oleobj/sample_with_lnk_to_calc.doc': (True, True, WORD, False,
                                                   True, False, False, False,
                                                   False, True, 0),
            'oleobj/embedded-simple-2007.ppt': (True, True, PPT, False, False,
                                                False, False, True, False,
                                                False, 0),
            'oleobj/sample_with_lnk_file.pps': (True, True, PPT, False, False,
                                                False, False, True, False,
                                                False, 0),
            'oleobj/embedded-simple-2007.pptx': (False,),
            'oleobj/embedded-simple-2007.ppsm': (False,),
            'oleobj/embedded-simple-2007.dotx': (False,),
            'oleobj/embedded-simple-2007.pptm': (False,),
            'oleobj/embedded-simple-2007.xlt': (True, True, EXCEL, False,
                                                False, False, True, False,
                                                False, False, 0),
            'oleobj/embedded-simple-2007.docx': (False,),
            'oleobj/embedded-simple-2007.potx': (False,),
            'oleobj/embedded-simple-2007.pot': (True, True, PPT, False, False,
                                                False, False, True, False,
                                                False, 0),
            'oleobj/embedded-simple-2007.xltm': (False,),
            'oleobj/embedded-simple-2007.potm': (False,),
            'encrypted/encrypted.xlsx': CRYPT,
            'encrypted/encrypted.docm': CRYPT,
            'encrypted/encrypted.docx': CRYPT,
            'encrypted/encrypted.pptm': CRYPT,
            'encrypted/encrypted.xlsb': CRYPT,
            'encrypted/encrypted.xls': (True, True, EXCEL, True, False, False,
                                        True, False, False, False, 0),
            'encrypted/encrypted.ppt': (True, False, 'unknown', True, False,
                                        False, False, True, False, False, 0),
            'encrypted/encrypted.pptx': CRYPT,
            'encrypted/encrypted.xlsm': CRYPT,
            'encrypted/encrypted.doc': (True, True, WORD, True, True, False,
                                        False, False, False, False, 0),
            'msodde/harmless-clean.docm': (False,),
            'msodde/dde-in-csv.csv': (False,),
            'msodde/dde-test-from-office2013-utf_16le-korean.doc':
                (True, True, WORD, False, True, False, False, False, False,
                 False, 0),
            'msodde/harmless-clean.doc': (True, True, WORD, False, True, False,
                                          False, False, False, False, 0),
            'msodde/dde-test.docm': (False,),
            'msodde/dde-test.xlsb': (False,),
            'msodde/dde-test.xlsm': (False,),
            'msodde/dde-test.docx': (False,),
            'msodde/dde-test.xlsx': (False,),
            'msodde/dde-test-from-office2003.doc': (True, True, WORD, False,
                                                    True, False, False, False,
                                                    False, False, 0),
            'msodde/dde-test-from-office2016.doc': (True, True, WORD, False,
                                                    True, False, False, False,
                                                    False, False, 0),
            'msodde/harmless-clean.docx': (False,),
            'oleform/oleform-PR314.docm': (False,),
            'basic/encrypted.docx': CRYPT,
            'oleobj/external_link/sample_with_external_link_to_doc.docx': (False,),
            'oleobj/external_link/sample_with_external_link_to_doc.xlsb': (False,),
            'oleobj/external_link/sample_with_external_link_to_doc.dotm': (False,),
            'oleobj/external_link/sample_with_external_link_to_doc.xlsm': (False,),
            'oleobj/external_link/sample_with_external_link_to_doc.pptx': (False,),
            'oleobj/external_link/sample_with_external_link_to_doc.dotx': (False,),
            'oleobj/external_link/sample_with_external_link_to_doc.docm': (False,),
            'oleobj/external_link/sample_with_external_link_to_doc.potm': (False,),
            'oleobj/external_link/sample_with_external_link_to_doc.xlsx': (False,),
            'oleobj/external_link/sample_with_external_link_to_doc.potx': (False,),
            'oleobj/external_link/sample_with_external_link_to_doc.ppsm': (False,),
            'oleobj/external_link/sample_with_external_link_to_doc.pptm': (False,),
            'oleobj/external_link/sample_with_external_link_to_doc.ppsx': (False,),
            'encrypted/autostart-encrypt-standardpassword.xlsm':
                (True, False, 'unknown', True, False, False, False, False, False, False, 0),
            'encrypted/autostart-encrypt-standardpassword.xls':
                (True, True, EXCEL, True, False, True, True, False, False, False, 0),
            'encrypted/dde-test-encrypt-standardpassword.xlsx':
                (True, False, 'unknown', True, False, False, False, False, False, False, 0),
            'encrypted/dde-test-encrypt-standardpassword.xlsm':
                (True, False, 'unknown', True, False, False, False, False, False, False, 0),
            'encrypted/autostart-encrypt-standardpassword.xlsb':
                (True, False, 'unknown', True, False, False, False, False, False, False, 0),
            'encrypted/dde-test-encrypt-standardpassword.xls':
                (True, True, EXCEL, True, False, False, True, False, False, False, 0),
            'encrypted/dde-test-encrypt-standardpassword.xlsb':
                (True, False, 'unknown', True, False, False, False, False, False, False, 0),
        }

        indicator_names = []
        for base_dir, _, files in os.walk(DATA_BASE_DIR):
            for filename in files:
                full_path = join(base_dir, filename)
                name = relpath(full_path, DATA_BASE_DIR)
                values = tuple(indicator.value for indicator in
                               oleid.OleID(full_path).check())
                if len(indicator_names) < 2:   # not initialized with ole yet
                    indicator_names = tuple(indicator.name for indicator in
                                            oleid.OleID(full_path).check())
                suffix = splitext(filename)[1]
                if suffix in NON_OLE_SUFFIXES:
                    self.assertEqual(values, NON_OLE_VALUES,
                                     msg='For non-ole file {} expected {}, '
                                         'not {}'.format(name, NON_OLE_VALUES,
                                                         values))
                    continue
                try:
                    self.assertEqual(values, OLE_VALUES[name],
                                     msg='Wrong detail values for {}:\n'
                                         '  Names  {}\n  Found  {}\n  Expect {}'
                                         .format(name, indicator_names, values,
                                                 OLE_VALUES[name]))
                except KeyError:
                    print('Should add oleid output for {} to {} ({})'
                          .format(name, __name__, values))


# just in case somebody calls this file as a script
if __name__ == '__main__':
    unittest.main()
