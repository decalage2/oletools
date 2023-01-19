""" Test ZipSubFile

Checks that ZipSubFile behaves just like a regular file-like object, just with
a few less allowed operations.
"""

import unittest
from tempfile import mkstemp, TemporaryFile
import os
from zipfile import ZipFile

from oletools.ooxml import ZipSubFile


# flag to get more output to facilitate search for errors
DEBUG = False

# name of a temporary .zip file on the system
ZIP_TEMP_FILE = ''

# name of a file inside the temporary zip file
FILE_NAME = 'test.txt'

# contents of that file
FILE_CONTENTS = b'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ'


def setUpModule():
    """ Called once before the first test; creates a temp zip file """
    global ZIP_TEMP_FILE
    handle, ZIP_TEMP_FILE = mkstemp(suffix='.zip',
                                    prefix='oletools-test-ZipSubFile-')
    os.close(handle)

    with ZipFile(ZIP_TEMP_FILE, 'w') as writer:
        writer.writestr(FILE_NAME, FILE_CONTENTS)
    if DEBUG:
        print('Created zip file ' + ZIP_TEMP_FILE)


def tearDownModule():
    """ Called once after last test; removes the temp zip file """
    if ZIP_TEMP_FILE and os.path.isfile(ZIP_TEMP_FILE):
        if DEBUG:
            print('leaving temp zip file {0} for inspection'
                  .format(ZIP_TEMP_FILE))
        else:
            os.unlink(ZIP_TEMP_FILE)
    elif DEBUG:
        print('WARNING: zip temp file apparently not created')


class TestZipSubFile(unittest.TestCase):
    """ Tests ZipSubFile """

    def setUp(self):
        self.zipper = ZipFile(ZIP_TEMP_FILE)
        self.subfile = ZipSubFile(self.zipper, FILE_NAME)
        self.subfile.open()

        # create a file in memory for comparison
        self.compare = TemporaryFile(prefix='oletools-test-ZipSubFile-',
                                     suffix='.bin')
        self.compare.write(FILE_CONTENTS)
        self.compare.seek(0)   # re-position to start

        self.assertEqual(self.subfile.tell(), 0)
        self.assertEqual(self.compare.tell(), 0)
        if DEBUG:
            print('created comparison file {0!r} in memory'
                  .format(self.compare.name))

    def tearDown(self):
        self.compare.close()
        self.subfile.close()
        self.zipper.close()
        if DEBUG:
            print('\nall files closed')

    def test_read(self):
        """ test reading """
        # read from start
        self.assertEqual(self.subfile.read(4), self.compare.read(4))
        self.assertEqual(self.subfile.tell(), self.compare.tell())

        # read a bit more
        self.assertEqual(self.subfile.read(4), self.compare.read(4))
        self.assertEqual(self.subfile.tell(), self.compare.tell())

        # create difference
        self.subfile.read(1)
        self.assertNotEqual(self.subfile.read(4), self.compare.read(4))
        self.compare.read(1)
        self.assertEqual(self.subfile.tell(), self.compare.tell())

        # read all the rest
        self.assertEqual(self.subfile.read(), self.compare.read())
        self.assertEqual(self.subfile.tell(), self.compare.tell())

    def test_seek_forward(self):
        """ test seeking forward """
        self.subfile.seek(10)
        self.compare.seek(10)
        self.assertEqual(self.subfile.read(1), self.compare.read(1))
        self.assertEqual(self.subfile.tell(), self.compare.tell())

        # seek 2 forward
        self.subfile.seek(2, os.SEEK_CUR)
        self.compare.seek(2, os.SEEK_CUR)
        self.assertEqual(self.subfile.read(1), self.compare.read(1))
        self.assertEqual(self.subfile.tell(), self.compare.tell())

        # seek backward (only implemented case: back to start)
        self.subfile.seek(-1 * self.subfile.tell(), os.SEEK_CUR)
        self.compare.seek(-1 * self.compare.tell(), os.SEEK_CUR)
        self.assertEqual(self.subfile.read(1), self.compare.read(1))
        self.assertEqual(self.subfile.tell(), self.compare.tell())

        # seek to end
        self.subfile.seek(0, os.SEEK_END)
        self.compare.seek(0, os.SEEK_END)
        self.assertEqual(self.subfile.tell(), self.compare.tell())

        # seek back to start
        self.subfile.seek(0)
        self.compare.seek(0)
        self.assertEqual(self.subfile.tell(), self.compare.tell())
        self.assertEqual(self.subfile.tell(), 0)

    def test_check_size(self):
        """ test usual size check: seek to end, tell, seek to start """
        # seek to end
        self.subfile.seek(0, os.SEEK_END)
        self.assertEqual(self.subfile.tell(), len(FILE_CONTENTS))

        # seek back to start
        self.subfile.seek(0)

        # read first few bytes
        self.assertEqual(self.subfile.read(10), FILE_CONTENTS[:10])

    def test_error_read(self):
        """ test correct behaviour if read beyond end (no exception) """
        self.subfile.seek(0, os.SEEK_END)
        self.compare.seek(0, os.SEEK_END)

        self.assertEqual(self.compare.read(10), self.subfile.read(10))
        self.assertEqual(self.compare.tell(), self.subfile.tell())

        self.subfile.seek(0)
        self.compare.seek(0)
        self.subfile.seek(len(FILE_CONTENTS) - 1)
        self.compare.seek(len(FILE_CONTENTS) - 1)
        self.assertEqual(self.compare.read(10), self.subfile.read(10))
        self.assertEqual(self.compare.tell(), self.subfile.tell())

    def test_error_seek(self):
        """ test correct behaviour if seek beyond end (no exception) """
        self.subfile.seek(len(FILE_CONTENTS) + 10)
        self.compare.seek(len(FILE_CONTENTS) + 10)
        # subfile.tell() gives len(FILE_CONTENTS),
        # compare.tell() gives len(FILE_CONTENTS) + 10,
        #self.assertEqual(self.subfile.tell(), self.compare.tell())

# just in case somebody calls this file as a script
if __name__ == '__main__':
    unittest.main()
