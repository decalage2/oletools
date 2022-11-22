"""
Run all tools on all test data to check for regressions.
"""

import unittest

# Directory with test data, independent of current working directory
from tests.test_utils import loop_and_extract, call_and_capture, DATA_BASE_DIR


class TestOnAll(unittest.TestCase):
    """Run all tools on all test data."""

    def do_test(self, module, skip_list_arg=None):
        """Helper for the tests that does the actual work."""
        if skip_list_arg is None:
            skip_list = []
        else:
            skip_list = skip_list_arg

        for full_path, rel_path in loop_and_extract():
            if rel_path in skip_list:
                print('Run {0} on all test data: skip {1}'
                      .format(module, rel_path))
                continue

            output, return_code = call_and_capture(module, [full_path,],
                                                   accept_nonzero_exit=True)
            if return_code == 0:
                continue

            error = '{0} returned {1} for sample {2}' \
                    .format(module, return_code, rel_path)
            print(error)
            for line in output.splitlines():
                print(line.rstrip())
            self.fail(error)

    def test_olevba(self):
        """Run olevba on all test data"""
        skip_list = ('rtfobj/issue_185.rtf.zip',
                     'rtfobj/issue_251.rtf',
                     'msodde/RTF-Spec-1.7.rtf',
                     'basic/encrypted.docx',
                     'encrypted/encrypted.doc',
                     'encrypted/encrypted.docm',
                     'encrypted/encrypted.docx',
                     'encrypted/encrypted.ppt',
                     'encrypted/encrypted.pptm',
                     'encrypted/encrypted.pptx',
                     'encrypted/encrypted.xls',
                     'encrypted/encrypted.xlsm',
                     'encrypted/encrypted.xlsx',
                     'encrypted/encrypted.xlsb',
                     )
        self.do_test('olevba', skip_list)

    # todo: add all the others as well

# just in case somebody calls this file as a script
if __name__ == '__main__':
    unittest.main()
