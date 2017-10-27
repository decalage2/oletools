""" Test validity of json output

Some scripts have a json output flag. Verify that at default log levels output
can be captured as-is and parsed by a json parser -- at least if the scripts
return 0
"""

import unittest
import sys
import json
import os
from os.path import join
from oletools import msodde
from tests.test_utils import OutputCapture, DATA_BASE_DIR

if sys.version_info[0] <= 2:
    from oletools import olevba
else:
    from oletools import olevba3 as olevba


class TestValidJson(unittest.TestCase):
    """ Ensure that script output is valid json (if return code is 0) """

    def iter_test_files(self):
        """ Iterate over all test files in DATA_BASE_DIR """
        for dirpath, _, filenames in os.walk(DATA_BASE_DIR):
            for filename in filenames:
                yield join(dirpath, filename)

    def run_and_parse(self, program, args, print_output=False):
        """ run single program with single file and parse output """
        return_code = None
        with OutputCapture() as capturer:       # capture stdout
            try:
                return_code = program(args)
            except Exception:
                return_code = 1   # would result in non-zero exit
            except SystemExit as se:
                return_code = se.code or 0   # se.code can be None
        if return_code is not 0:
            if print_output:
                print('Command failed ({0}) -- not parsing output'
                      .format(return_code))
            return    # no need to test

        self.assertNotEqual(return_code, None,
                            msg='self-test fail: return_code not set')

        # now test output
        if print_output:
            print(capturer.buffer.getvalue())
        capturer.buffer.seek(0)    # re-set position in file-like stream
        try:
            json_data = json.load(capturer.buffer)
        except ValueError:
            self.fail('Invalid json:\n' + capturer.buffer.getvalue())
        self.assertNotEqual(len(json_data), 0, msg='Output was empty')

    def run_all_files(self, program, args_without_filename, print_output=False):
        """ run test for a single program over all test files """
        n_files = 0
        for testfile in self.iter_test_files():   # loop over all input
            args = args_without_filename + [testfile, ]
            self.run_and_parse(program, args, print_output)
            n_files += 1
        self.assertNotEqual(n_files, 0,
                            msg='self-test fail: No test files found')

    def test_msodde(self):
        """ Test msodde.py """
        self.run_all_files(msodde.main, ['-j', ])

    @unittest.skip('olevba needs patching to accept custom cmd line args')
    def test_olevba(self):
        """ Test olevba.py with default args """
        self.run_all_files(olevba.main, ['-j', ])

    @unittest.skip('olevba needs patching to accept custom cmd line args')
    def test_olevba_analysis(self):
        """ Test olevba.py with default args """
        self.run_all_files(olevba.main, ['-j', '-a', ])

    @unittest.skip('olevba needs patching to accept custom cmd line args')
    def test_olevba_recurse(self):
        """ Test olevba.py with default args """
        self.run_and_parse(olevba.main, ['-j', '-r', DATA_DIR], True)


# just in case somebody calls this file as a script
if __name__ == '__main__':
    unittest.main()
