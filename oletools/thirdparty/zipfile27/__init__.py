# Excerpt from the zipfile module from Python 2.7, to enable is_zipfile
# to check any file object (e.g. in memory), for Python 2.6.
# is_zipfile in Python 2.6 can only check files on disk.

# This code from Python 2.7 was not modified.

# 2016-09-06 v0.01 PL: - first version


from zipfile import _EndRecData

def _check_zipfile(fp):
    try:
        if _EndRecData(fp):
            return True         # file has correct magic number
    except IOError:
        pass
    return False

def is_zipfile(filename):
    """Quickly see if a file is a ZIP file by checking the magic number.

    The filename argument may be a file or file-like object too.
    """
    result = False
    try:
        if hasattr(filename, "read"):
            result = _check_zipfile(fp=filename)
        else:
            with open(filename, "rb") as fp:
                result = _check_zipfile(fp)
    except IOError:
        pass
    return result

