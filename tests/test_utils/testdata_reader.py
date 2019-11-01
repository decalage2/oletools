import os, sys, zipfile
from os.path import dirname, abspath, normpath, join
from . import DATA_BASE_DIR

ENCRYPTED_FILES_PASSWORD='infected-test'

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
    with open(get_path_from_root(relative_path), 'rb') as file_handle:
        return file_handle.read()

def read_encrypted(relative_path, filename=None):
    z = zipfile.ZipFile(get_path_from_root(relative_path))

    if filename == None:
        contents = z.read(z.namelist()[0], pwd=ENCRYPTED_FILES_PASSWORD)
    else:
        contents = z.read(filename, pwd=ENCRYPTED_FILES_PASSWORD)

    z.close()
    return contents

def get_path_from_root(relative_path):
    return join(DATA_BASE_DIR, relative_path)