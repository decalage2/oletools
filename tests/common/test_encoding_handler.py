"""Test common.ensure_stdout_handles_unicode"""

from __future__ import print_function

import unittest
import sys
from subprocess import check_call, CalledProcessError
from tempfile import mkstemp
import os
from os.path import isfile
from contextlib import contextmanager

FILE_TEXT = u'The unicode check mark is \u2713.\n'

@contextmanager
def temp_file(just_name=True):
    """Context manager that creates temp file and deletes it in the end"""
    tmp_descriptor = None
    tmp_name = None
    tmp_handle = None
    try:
        tmp_descriptor, tmp_name = mkstemp()

        # we create our own file handle since we want to be able to close the
        # file and open it again for reading.
        # We keep the os-level descriptor open so file name is still reserved
        # for us
        if just_name:
            yield tmp_name
        else:
            tmp_handle = open(tmp_name, 'wb')
            yield tmp_handle, tmp_name
    except Exception:
        raise
    finally:
        if tmp_descriptor is not None:
            os.close(tmp_descriptor)
        if tmp_handle is not None:
            tmp_handle.close()
        if tmp_name is not None and isfile(tmp_name):
            os.unlink(tmp_name)


class TestEncodingHandler(unittest.TestCase):
    """Tests replacing stdout encoding in various scenarios"""

    def test_print(self):
        """Test regular unicode output not raise error"""
        check_call('{python} {this_file} print'.format(python=sys.executable,
                                                      this_file=__file__),
                   shell=True)

    def test_print_redirect(self):
        """
        Test redirection of unicode output to files does not raise error

        TODO: test this on non-linux OSs
        """
        with temp_file() as tmp_file:
            check_call('{python} {this_file} print > {tmp_file}'
                       .format(python=sys.executable, this_file=__file__,
                               tmp_file=tmp_file),
                       shell=True)

    @unittest.skipIf(not sys.platform.startswith('linux'),
                     'Only tested on linux sofar')
    def test_print_no_lang(self):
        """
        Test redirection of unicode output to files does not raise error

        TODO: Adapt this for other OSs; for win create batch script
        """
        check_call('LANG=C {python} {this_file} print'
                   .format(python=sys.executable, this_file=__file__),
                   shell=True)

    def test_uopen(self):
        """Test that uopen in a nice environment is ok"""
        with temp_file(False) as (tmp_handle, tmp_file):
            tmp_handle.write(FILE_TEXT.encode('utf8'))
            tmp_handle.close()

            try:
                check_call('{python} {this_file} read {tmp_file}'
                           .format(python=sys.executable, this_file=__file__,
                                   tmp_file=tmp_file),
                           shell=True)
            except CalledProcessError as cpe:
                self.fail(cpe.output)

    def test_uopen_redirect(self):
        """
        Test redirection of unicode output to files does not raise error

        TODO: test this on non-linux OSs
        """
        with temp_file(False) as (tmp_handle, tmp_file):
            tmp_handle.write(FILE_TEXT.encode('utf8'))
            tmp_handle.close()

            with temp_file() as redirect_file:
                try:
                    check_call(
                        '{python} {this_file} read {tmp_file} >{redirect_file}'
                        .format(python=sys.executable, this_file=__file__,
                        tmp_file=tmp_file, redirect_file=redirect_file),
                        shell=True)
                except CalledProcessError as cpe:
                    self.fail(cpe.output)

    @unittest.skipIf(not sys.platform.startswith('linux'),
                     'Only tested on linux sofar')
    def test_uopen_no_lang(self):
        """
        Test that uopen in a C-LANG environment is ok

        TODO: Adapt this for other OSs; for win create batch script
        """
        with temp_file(False) as (tmp_handle, tmp_file):
            tmp_handle.write(FILE_TEXT.encode('utf8'))
            tmp_handle.close()

            try:
                check_call('LANG=C {python} {this_file} read {tmp_file}'
                           .format(python=sys.executable, this_file=__file__,
                                   tmp_file=tmp_file),
                           shell=True)
            except CalledProcessError as cpe:
                self.fail(cpe.output)


def run_read(filename):
    """This is called from test_uopen* tests as script. Reads text, compares"""
    from oletools.common.io_encoding import uopen
    # open file
    with uopen(filename, 'rt') as reader:
        # a few tests
        if reader.closed:
            raise ValueError('handle is closed!')
        if reader.name != filename:
            raise ValueError('Wrong filename {}'.format(reader.name))
        if reader.isatty():
            raise ValueError('Reader is a tty!')
        if reader.tell() != 0:
            raise ValueError('Reader.tell is not 0 at beginning')

        # read text
        text = reader.read()

    # a few more tests
    if not reader.closed:
        raise ValueError('Reader is not closed outside context')
    if reader.name != filename:
        raise ValueError('Wrong filename {} after context'.format(reader.name))
    # the following test raises an exception because reader is closed, so isatty cannot be called:
    # if reader.isatty():
    #     raise ValueError('Reader has become a tty!')

    # compare text
    if sys.version_info.major <= 2:      # in python2 get encoded byte string
        expect = FILE_TEXT.encode('utf8')
    else:                                # python3: should get real unicode
        expect = FILE_TEXT
    if text != expect:
        raise ValueError('Wrong contents: {!r} != {!r}'
                         .format(text, expect))
    return 0


def run_print():
    """This is called from test_read* tests as script. Prints & logs unicode"""
    from oletools.common.io_encoding import ensure_stdout_handles_unicode
    from oletools.common.log_helper import log_helper
    ensure_stdout_handles_unicode()
    print(u'Check: \u2713')    # print check mark

    # check logging as well
    logger = log_helper.get_or_create_silent_logger('test_encoding_handler')
    log_helper.enable_logging(False, 'debug', stream=sys.stdout)
    logger.info(u'Check: \u2713')
    return 0


# tests call this file as script
if __name__ == '__main__':
    if len(sys.argv) < 2:
        sys.exit(unittest.main())

    # hack required to import common from parent dir, not system-wide one
    # (usually unittest seems to do that for us)
    from os.path import abspath, dirname, join
    ole_base = dirname(dirname(dirname(abspath(__file__))))
    sys.path.insert(0, ole_base)

    if sys.argv[1] == 'print':
        if len(sys.argv) > 2:
            print('Expect no arg for "print"', file=sys.stderr)
            sys.exit(2)
        sys.exit(run_print())
    elif sys.argv[1] == 'read':
        if len(sys.argv) != 3:
            print('Expect single arg for "read"', file=sys.stderr)
            sys.exit(2)
        sys.exit(run_read(sys.argv[2]))
    else:
        print('Unexpected argument: {}'.format(sys.argv[1]), file=sys.stderr)
        sys.exit(2)
