""" Test oleform basic functionality """

import unittest
from os.path import join
import sys

# Directory with test data, independent of current working directory
from tests.test_utils import DATA_BASE_DIR

if sys.version_info[0] <= 2:
    from oletools.olevba import VBA_Parser
else:
    from oletools.olevba3 import VBA_Parser

SAMPLES = [('oleform-PR314.docm', [('word/vbaProject.bin', u'UserFormTEST1', {'caption': 'Label1-test', 'control_tip_text': None, 'name': 'Label1', 'value': None, 'tag': 'l\x18sdf', 'ClsidCacheIndex': 21, 'id': 1, 'tabindex': 0}), ('word/vbaProject.bin', u'UserFormTEST1', {'caption': None, 'control_tip_text': None, 'name': 'TextBox1', 'value': 'heyhey', 'tag': '', 'ClsidCacheIndex': 23, 'id': 2, 'tabindex': 1}), ('word/vbaProject.bin', u'UserFormTEST1', {'caption': None, 'control_tip_text': None, 'name': 'ComboBox1', 'value': 'none dd', 'tag': '', 'ClsidCacheIndex': 25, 'id': 3, 'tabindex': 2}), ('word/vbaProject.bin', u'UserFormTEST1', {'caption': None, 'control_tip_text': None, 'name': 'CheckBox1', 'value': '1', 'tag': '', 'ClsidCacheIndex': 26, 'id': 5, 'tabindex': 4}), ('word/vbaProject.bin', u'UserFormTEST1', {'caption': None, 'control_tip_text': None, 'name': 'OptionButton1', 'value': '0', 'tag': '', 'ClsidCacheIndex': 27, 'id': 6, 'tabindex': 5}), ('word/vbaProject.bin', u'UserFormTEST1', {'caption': None, 'control_tip_text': None, 'name': 'ToggleButton1', 'value': '0', 'tag': '', 'ClsidCacheIndex': 28, 'id': 7, 'tabindex': 6}), ('word/vbaProject.bin', u'UserFormTEST1', {'caption': None, 'control_tip_text': None, 'name': 'Frame1', 'value': None, 'tag': '', 'ClsidCacheIndex': 14, 'id': 8, 'tabindex': 7}), ('word/vbaProject.bin', u'UserFormTEST1', {'caption': None, 'control_tip_text': None, 'name': 'TabStrip1', 'value': None, 'tag': '', 'ClsidCacheIndex': 18, 'id': 10, 'tabindex': 8}), ('word/vbaProject.bin', u'UserFormTEST1', {'caption': None, 'control_tip_text': None, 'name': 'CommandButton1', 'value': None, 'tag': '', 'ClsidCacheIndex': 17, 'id': 9, 'tabindex': 9}), ('word/vbaProject.bin', u'UserFormTEST1', {'caption': None, 'control_tip_text': None, 'name': 'MultiPage1', 'value': None, 'tag': '', 'ClsidCacheIndex': 57, 'id': 12, 'tabindex': 10}), ('word/vbaProject.bin', u'UserFormTEST1', {'caption': None, 'control_tip_text': None, 'name': 'ScrollBar1', 'value': None, 'tag': '', 'ClsidCacheIndex': 47, 'id': 16, 'tabindex': 11}), ('word/vbaProject.bin', u'UserFormTEST1', {'caption': None, 'control_tip_text': None, 'name': 'SpinButton1', 'value': None, 'tag': '', 'ClsidCacheIndex': 16, 'id': 17, 'tabindex': 12}), ('word/vbaProject.bin', u'UserFormTEST1', {'caption': None, 'control_tip_text': None, 'name': 'Image1', 'value': None, 'tag': '', 'ClsidCacheIndex': 12, 'id': 18, 'tabindex': 13}), ('word/vbaProject.bin', u'UserFormTEST1', {'caption': None, 'control_tip_text': None, 'name': 'ListBox1', 'value': '', 'tag': '', 'ClsidCacheIndex': 24, 'id': 4, 'tabindex': 3}), ('word/vbaProject.bin', u'UserFormTEST1/i08', {'caption': None, 'control_tip_text': None, 'name': 'TextBox2', 'value': 'abcd', 'tag': '', 'ClsidCacheIndex': 23, 'id': 20, 'tabindex': 0}), ('word/vbaProject.bin', u'UserFormTEST1/i12', {'caption': None, 'control_tip_text': None, 'name': '', 'value': None, 'tag': '', 'ClsidCacheIndex': 18, 'id': 13, 'tabindex': 2}), ('word/vbaProject.bin', u'UserFormTEST1/i12', {'caption': None, 'control_tip_text': None, 'name': 'Page1', 'value': None, 'tag': '', 'ClsidCacheIndex': 7, 'id': 14, 'tabindex': 0}), ('word/vbaProject.bin', u'UserFormTEST1/i12', {'caption': None, 'control_tip_text': None, 'name': 'Page2', 'value': None, 'tag': '', 'ClsidCacheIndex': 7, 'id': 15, 'tabindex': 1}), ('word/vbaProject.bin', u'UserFormTEST1/i12/i14', {'caption': None, 'control_tip_text': None, 'name': 'TextBox3', 'value': 'last one', 'tag': '', 'ClsidCacheIndex': 23, 'id': 24, 'tabindex': 0}), ('word/vbaProject.bin', u'UserFormTest2', {'caption': 'Label1', 'control_tip_text': None, 'name': 'Label1', 'value': None, 'tag': '', 'ClsidCacheIndex': 21, 'id': 1, 'tabindex': 0}), ('word/vbaProject.bin', u'UserFormTest2', {'caption': 'Label2', 'control_tip_text': None, 'name': 'Label2', 'value': None, 'tag': '', 'ClsidCacheIndex': 21, 'id': 2, 'tabindex': 1}), ('word/vbaProject.bin', u'UserFormTest2', {'caption': None, 'control_tip_text': None, 'name': 'TextBox1', 'value': '&\xe9"\'', 'tag': '', 'ClsidCacheIndex': 23, 'id': 3, 'tabindex': 2})])]

class TestOleForm(unittest.TestCase):

    def test_samples(self):
        if sys.version_info[0] > 2:
             # Unfortunately, olevba3 doesn't have extract_form_strings_extended
             return
        for sample, expected_result in SAMPLES:
            full_name = join(DATA_BASE_DIR, 'oleform', sample)
            parser = VBA_Parser(full_name)
            variables = list(parser.extract_form_strings_extended())
            self.assertEqual(variables, expected_result)


# just in case somebody calls this file as a script
if __name__ == '__main__':
    unittest.main()

