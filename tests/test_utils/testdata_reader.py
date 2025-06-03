"""
Helper functions to deal with zip-encrypted test files.

Some test samples alerted antivirus software when installing oletools. Those
samples were therefore "hidden" in encrypted zip-files. These functions help
using them.
"""

import os, sys, zipfile
from os.path import relpath, join, isfile, splitext
from contextlib import contextmanager
from tempfile import mkstemp, TemporaryDirectory, NamedTemporaryFile

from . import DATA_BASE_DIR

# Passwort used to encrypt problematic test samples inside a zip container
ENCRYPTED_FILES_PASSWORD='infected-test'

# import zipfile in a way compatible with all kinds of old python versions
if sys.version_info[0] <= 2:
    # Python 2.x
    if sys.version_info[1] <= 6:
        # Python 2.6
        # use is_zipfile backported from Python 2.7:
        from thirdparty.zipfile27 import is_zipfile
    else:
        # Python 2.7
        from zipfile import is_zipfile
else:
    # Python 3.x+
    from zipfile import is_zipfile
    ENCRYPTED_FILES_PASSWORD = ENCRYPTED_FILES_PASSWORD.encode()


def read(relative_path):
    """
    Return contents of unencrypted file inside test data dir.

    see also: `read_encrypted`.
    """
    with open(get_path_from_root(relative_path), 'rb') as file_handle:
        return file_handle.read()


def read_encrypted(relative_path, filename=None):
    """
    Return contents of encrypted file inside test data dir.

    see also: `read`.
    """
    z = zipfile.ZipFile(get_path_from_root(relative_path))

    if filename == None:
        contents = z.read(z.namelist()[0], pwd=ENCRYPTED_FILES_PASSWORD)
    else:
        contents = z.read(filename, pwd=ENCRYPTED_FILES_PASSWORD)

    z.close()
    return contents


def get_path_from_root(relative_path):
    """Convert path relative to test data base dir to an absolute path."""
    return join(DATA_BASE_DIR, relative_path)


def loop_over_files(subdir=''):
    """
    Find all files, decrypting problematic files on the fly.

    Does a `os.walk` through all test data or the given subdir and yields a
    2-tuple for each sample: the path to the file relative to `DATA_BASE_DIR`
    and the contents of the file, with the file being unzipped first if it ends
    with .zip.

    See also: :py:meth:`loop_and_extract`

    :param str subdir: Optional subdir of test data dir that caller is interested in
    """
    for base_dir, _, files in os.walk(join(DATA_BASE_DIR, subdir)):
        for filename in files:
            relative_path = relpath(join(base_dir, filename), DATA_BASE_DIR)
            if filename.endswith('.zip'):
                yield relative_path, read_encrypted(relative_path)
            else:
                yield relative_path, read(relative_path)


def loop_and_extract(subdir=''):
    """
    Find all files, decrypting them to tempdir if necessary.

    Does a `os.walk` through all test data or the given subdir and yields
    the absolute path for each sample, which is either its original location
    in `DATA_BASE_DIR` or in a temporary directory if it had to be decrypted.

    The temp dir and files inside it are always deleted right after usage.

    See also: :py:meth:`loop_over_files`

    :param str subdir: Optional subdir of test data dir that caller is interested in
    """
    with TemporaryDirectory(prefix='oletools-test-') as temp_dir:
        for base_dir, _, files in os.walk(join(DATA_BASE_DIR, subdir)):
            for filename in files:
                full_path = join(base_dir, filename)
                if filename.endswith('.zip'):
                    # remove the ".zip" and split the rest into actual name and extension
                    actual_name, actual_extn = splitext(splitext(filename)[0])

                    with zipfile.ZipFile(full_path, 'r') as zip_file:
                        # create a temp file that has a proper file name and is deleted on closing
                        with NamedTemporaryFile(dir=temp_dir, prefix=actual_name, suffix=actual_extn) \
                                as temp_file:
                            # our test samples are not big, so we can read the whole thing at once
                            temp_file.write(zip_file.read(zip_file.namelist()[0],
                                                          pwd=ENCRYPTED_FILES_PASSWORD))
                            temp_file.flush()
                            yield temp_file.name
                else:
                    yield full_path


@contextmanager
def decrypt_sample(relpath):
    """
    Decrypt test sample, save to tempfile, yield temp file name.

    Use as context-manager, deletes tempfile after use.

    If sample is not encrypted at all (filename does not end in '.zip'),
    yields absolute path to sample itself, so can apply this code also
    to non-encrypted samples.

    Code based on test_encoding_handler.temp_file().

    Note: this causes problems if running with PyPy on Windows. The `unlink`
          fails because the file is "still being used by another process".

    :param relpath: path inside `DATA_BASE_DIR`, should end in '.zip'
    :return: absolute path name to decrypted sample.
    """
    if not relpath.endswith('.zip'):
        yield get_path_from_root(relpath)
    else:
        tmp_descriptor = None
        tmp_name = None
        try:
            tmp_descriptor, tmp_name = mkstemp(text=False)
            with zipfile.ZipFile(get_path_from_root(relpath), 'r') as unzipper:
                # no need to iterate over blobs, our test files are all small
                os.write(tmp_descriptor, unzipper.read(unzipper.namelist()[0],
                                                       pwd=ENCRYPTED_FILES_PASSWORD))
            os.close(tmp_descriptor)
            tmp_descriptor = None
            yield tmp_name
        except Exception:
            raise
        finally:
            if tmp_descriptor is not None:
                os.close(tmp_descriptor)
            if tmp_name is not None and isfile(tmp_name):
                os.unlink(tmp_name)