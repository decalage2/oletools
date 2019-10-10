""" Test oleobj basic functionality """

import unittest
from tempfile import mkdtemp
from shutil import rmtree
from os.path import join, isfile
from hashlib import md5
from glob import glob

# Directory with test data, independent of current working directory
from tests.test_utils import DATA_BASE_DIR, call_and_capture
from oletools import oleobj
from oletools.common.io_encoding import ensure_stdout_handles_unicode


#: provide some more info to find errors
DEBUG = False


# test samples in test-data/oleobj: filename, embedded file name, embedded md5
SAMPLES = (
    ('sample_with_calc_embedded.doc', 'calc.exe',
     '40e85286357723f326980a3b30f84e4f'),
    ('sample_with_lnk_file.doc', 'calc.lnk',
     '6aedb1a876d4ad5236f1fbbbeb7274f3'),
    ('sample_with_lnk_file.pps', 'calc.lnk',
     '6aedb1a876d4ad5236f1fbbbeb7274f3'),
    ('sample_with_lnk_file.ppt', 'calc.lnk',
     '6aedb1a876d4ad5236f1fbbbeb7274f3'),
    ('embedded-unicode.doc', '_nic_de-___________.txt',
     '264397735b6f09039ba0adf0dc9fb942'),
    ('embedded-unicode-2007.docx', '_nic_de-___________.txt',
     '264397735b6f09039ba0adf0dc9fb942'),
)
SAMPLES += tuple(
    ('embedded-simple-2007.' + extn, 'simple-text-file.txt',
     'bd5c063a5a43f67b3c50dc7b0f1195af')
    for extn in ('doc', 'dot', 'docx', 'docm', 'dotx', 'dotm')
)
SAMPLES += tuple(
    ('embedded-simple-2007.' + extn, 'simple-text-file.txt',
     'ab8c65e4c0fc51739aa66ca5888265b4')
    for extn in ('xls', 'xlsx', 'xlsb', 'xlsm', 'xla', 'xlam', 'xlt', 'xltm',
                 'xltx', 'ppt', 'pptx', 'pptm', 'pps', 'ppsx', 'ppsm', 'pot',
                 'potx', 'potm', 'ods', 'odp')
)
SAMPLES += (('embedded-simple-2007.odt', 'simple-text-file.txt',
     'bd5c063a5a43f67b3c50dc7b0f1195af'), )


def calc_md5(filename):
    """ calc md5sum of given file in temp_dir """
    chunk_size = 4096
    hasher = md5()
    with open(filename, 'rb') as handle:
        buf = handle.read(chunk_size)
        while buf:
            hasher.update(buf)
            buf = handle.read(chunk_size)
    return hasher.hexdigest()


def preread_file(args):
    """helper for TestOleObj.test_non_streamed: preread + call process_file"""
    ensure_stdout_handles_unicode()   # usually, main() call this
    ignore_arg, output_dir, filename = args
    if ignore_arg != '-d':
        raise ValueError('ignore_arg not as expected!')
    with open(filename, 'rb') as file_handle:
        data = file_handle.read()
    err_stream, err_dumping, did_dump = \
        oleobj.process_file(filename, data, output_dir=output_dir)
    if did_dump and not err_stream and not err_dumping:
        return oleobj.RETURN_DID_DUMP
    else:
        return oleobj.RETURN_NO_DUMP   # just anything else


class TestOleObj(unittest.TestCase):
    """ Tests oleobj basic feature """

    def setUp(self):
        """ fixture start: create temp dir """
        self.temp_dir = mkdtemp(prefix='oletools-oleobj-')
        self.did_fail = False

    def tearDown(self):
        """ fixture end: remove temp dir """
        if self.did_fail and DEBUG:
            print('leaving temp dir {0} for inspection'.format(self.temp_dir))
        elif self.temp_dir:
            rmtree(self.temp_dir)

    def test_md5(self):
        """ test all files in oleobj test dir """
        self.do_test_md5(['-d', self.temp_dir])

    def test_md5_args(self):
        """
        test that oleobj can be called with -i and -v

        This is how ripOLE used to be often called (e.g. by amavisd-new);
        ensure oleobj is a compatible replacement.
        """
        self.do_test_md5(['-d', self.temp_dir, '-v', '-i'])

    def test_no_output(self):
        """ test that oleobj does not find data where it should not """
        args = ['-d', self.temp_dir]
        for sample_name in ('sample_with_lnk_to_calc.doc',
                            'embedded-simple-2007.xml',
                            'embedded-simple-2007-as2003.xml'):
            full_name = join(DATA_BASE_DIR, 'oleobj', sample_name)
            output, ret_val = call_and_capture('oleobj', args + [full_name, ],
                                               accept_nonzero_exit=True)
            if glob(self.temp_dir + 'ole-object-*'):
                self.fail('found embedded data in {0}. Output:\n{1}'
                          .format(sample_name, output))
            self.assertEqual(ret_val, oleobj.RETURN_NO_DUMP,
                             msg='Wrong return value {} for {}. Output:\n{}'
                                 .format(ret_val, sample_name, output))

    def do_test_md5(self, args, test_fun=None, only_run_every=1):
        """ helper for test_md5 and test_md5_args """
        data_dir = join(DATA_BASE_DIR, 'oleobj')

        # name of sample, extension of embedded file, md5 hash of embedded file
        for sample_index, (sample_name, embedded_name, expect_hash) \
                in enumerate(SAMPLES):
            if sample_index % only_run_every != 0:
                continue
            args_with_path = args + [join(data_dir, sample_name), ]
            if test_fun is None:
                output, ret_val = call_and_capture('oleobj', args_with_path,
                                                   accept_nonzero_exit=True)
            else:
                ret_val = test_fun(args_with_path)
                output = '[output: see above]'
            self.assertEqual(ret_val, oleobj.RETURN_DID_DUMP,
                             msg='Wrong return value {} for {}. Output:\n{}'
                                 .format(ret_val, sample_name, output))
            expect_name = join(self.temp_dir,
                               sample_name + '_' + embedded_name)
            if not isfile(expect_name):
                self.did_fail = True
                self.fail('{0} not created from {1}. Output:\n{2}'
                          .format(expect_name, sample_name, output))
                continue
            md5_hash = calc_md5(expect_name)
            if md5_hash != expect_hash:
                self.did_fail = True
                self.fail('Wrong md5 {0} of {1} from {2}. Output:\n{3}'
                          .format(md5_hash, expect_name, sample_name, output))
                continue

    def test_non_streamed(self):
        """ Ensure old oleobj behaviour still works: pre-read whole file """
        return self.do_test_md5(['-d', self.temp_dir], test_fun=preread_file,
                                only_run_every=4)


# just in case somebody calls this file as a script
if __name__ == '__main__':
    unittest.main()
