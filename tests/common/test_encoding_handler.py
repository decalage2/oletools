"""Test common.ensure_stdout_handles_unicode"""

import unittest
import sys
from subprocess import check_call
from tempfile import mkstemp
import os
from os.path import isfile


class TestEncodingHandler(unittest.TestCase):
    """Tests replacing stdout encoding in various scenarios"""

    def test_base(self):
        """Test regular unicode output not raise error"""
        check_call('{python} {this_file} print'.format(python=sys.executable,
                                                      this_file=__file__),
                   shell=True)

    def test_redirect(self):
        """
        Test redirection of unicode output to files does not raise error

        TODO: test this on non-linux OSs
        """
        tmp_handle = None
        tmp_name = None
        try:
            tmp_handle, tmp_name = mkstemp()
            check_call('{python} {this_file} print > {tmp_file}'
                       .format(python=sys.executable, this_file=__file__,
                               tmp_file=tmp_file),
                       shell=True)
        except Exception:
            raise
        finally:
            if tmp_handle is not None:
                os.close(tmp_handle)
            if tmp_name is not None and isfile(tmp_name):
                os.unlink(tmp_name)

    @unittest.skipIf(not sys.platform.startswith('linux'),
                     'Only tested on linux sofar')
    def test_no_lang(self):
        """
        Test redirection of unicode output to files does not raise error

        TODO: Adapt this for other OSs; for win create batch script
        """
        check_call('LANG=C {python} {this_file} print'
                   .format(python=sys.executable, this_file=__file__),
                   shell=True)

def run_print():
    """This is called from test_read* tests as script. Prints & logs unicode"""
    # hack required to import common from parent dir, not system-wide one
    # (usually unittest seems to do that for us)
    from os.path import abspath, dirname, join
    ole_base = dirname(dirname(dirname(abspath(__file__))))
    sys.path.insert(0, ole_base)

    from oletools import common
    common.ensure_stdout_handles_unicode()
    print(u'\u2713')    # print check mark


# tests call this file as script
if __name__ == '__main__':
    if len(sys.argv) < 2:
        sys.exit(unittest.main())
    if sys.argv[1] == 'print':
        if len(sys.argv) > 2:
            print('Expect no arg for "print"', file=sys.stderr)
            sys.exit(2)
        sys.exit(run_print())
    else:
        print('Unexpected argument: {}'.format(sys.argv[1]), file=sys.stderr)
        sys.exit(2)
