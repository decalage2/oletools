import os
from os.path import dirname, abspath, normpath, join
from . import DATA_BASE_DIR


def read(relative_path):
    with open(join(DATA_BASE_DIR, relative_path), 'rb') as file_handle:
        return file_handle.read()
