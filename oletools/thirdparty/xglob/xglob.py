#! /usr/bin/env python2
"""
xglob

xglob is a python package to list files matching wildcards (*, ?, []),
extending the functionality of the glob module from the standard python
library (https://docs.python.org/2/library/glob.html).

Main features:
- recursive file listing (including subfolders)
- file listing within Zip archives
- helper function to open files specified as arguments, supporting files
  within zip archives encrypted with a password

Author: Philippe Lagadec - http://www.decalage.info
License: BSD, see source code or documentation

For more info and updates: http://www.decalage.info/xglob
"""

# LICENSE:
#
# xglob is copyright (c) 2013-2018, Philippe Lagadec (http://www.decalage.info)
# All rights reserved.
#
# Redistribution and use in source and binary forms, with or without modification,
# are permitted provided that the following conditions are met:
#
#  * Redistributions of source code must retain the above copyright notice, this
#    list of conditions and the following disclaimer.
#  * Redistributions in binary form must reproduce the above copyright notice,
#    this list of conditions and the following disclaimer in the documentation
#    and/or other materials provided with the distribution.
#
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
# ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
# WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
# FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
# DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
# SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
# CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
# OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
# OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.


#------------------------------------------------------------------------------
# CHANGELOG:
# 2013-12-04 v0.01 PL: - scan several files from command line args
# 2014-01-14 v0.02 PL: - added riglob, ziglob
# 2014-12-26 v0.03 PL: - moved code from balbuzard into a separate package
# 2015-01-03 v0.04 PL: - fixed issues in iter_files + yield container name
# 2016-02-24 v0.05 PL: - do not stop on exceptions, return them as data
#                      - fixed issue when using wildcards with empty path
# 2016-04-28 v0.06 CH: - improved handling of non-existing files
#                        (by Christian Herdtweck)
# 2018-12-08 v0.07 PL: - fixed issue #373, zip password must be bytes

__version__ = '0.07'


#=== IMPORTS =================================================================

import os, fnmatch, glob, zipfile

#=== EXCEPTIONS ==============================================================

class PathNotFoundException(Exception):
    """ raised if given a fixed file/dir (not a glob) that does not exist """
    def __init__(self, path):
        super(PathNotFoundException, self).__init__(
            'Given path does not exist: %r' % path)


#=== FUNCTIONS ===============================================================

# recursive glob function to find files in any subfolder:
# inspired by http://stackoverflow.com/questions/14798220/how-can-i-search-sub-folders-using-glob-glob-module-in-python
def rglob (path, pattern='*.*'):
    """
    Recursive glob:
    similar to glob.glob, but finds files recursively in all subfolders of path.
    path: root directory where to search files
    pattern: pattern for filenames, using wildcards, e.g. *.txt
    """
    #TODO: more compatible API with glob: use single param, split path from pattern
    return [os.path.join(dirpath, f)
        for dirpath, dirnames, files in os.walk(path)
        for f in fnmatch.filter(files, pattern)]


def riglob (pathname):
    """
    Recursive iglob:
    similar to glob.iglob, but finds files recursively in all subfolders of path.
    pathname: root directory where to search files followed by pattern for
    filenames, using wildcards, e.g. *.txt
    """
    path, filespec = os.path.split(pathname)
    # fix path if empty:
    if path == '':
        path = '.'
    # print 'riglob: path=%r, filespec=%r' % (path, filespec)
    for dirpath, dirnames, files in os.walk(path):
        for f in fnmatch.filter(files, filespec):
            yield os.path.join(dirpath, f)


def ziglob (zipfileobj, pathname):
    """
    iglob in a zip:
    similar to glob.iglob, but finds files within a zip archive.
    - zipfileobj: zipfile.ZipFile object
    - pathname: root directory where to search files followed by pattern for
    filenames, using wildcards, e.g. *.txt
    """
    files = zipfileobj.namelist()
    #for f in files: print f
    for f in fnmatch.filter(files, pathname):
        yield f


def iter_files(files, recursive=False, zip_password=None, zip_fname='*'):
    """
    Open each file provided as argument:
    - files is a list of arguments
    - if zip_password is None, each file is listed without reading its content.
      Wilcards are supported.
    - if not, then each file is opened as a zip archive with the provided password
    - then files matching zip_fname are opened from the zip archive

    Iterator: yields (container, filename, data) for each file. If zip_password is None, then
    only the filename is returned, container and data=None. Otherwise container is the
    filename of the container (zip file), and data is the file content (or an exception).
    If a given filename is not a glob and does not exist, the triplet
    (None, filename, PathNotFoundException) is yielded. (Globs matching nothing
    do not trigger exceptions)
    """
    #TODO: catch exceptions and yield them for the caller (no file found, file is not zip, wrong password, etc)
    #TODO: use logging instead of printing
    #TODO: split in two simpler functions, the caller knows if it's a zip or not
    # print 'iter_files: files=%r, recursive=%s' % (files, recursive)
    # choose recursive or non-recursive iglob:
    if recursive:
        iglob = riglob
    else:
        iglob = glob.iglob
    for filespec in files:
        if not is_glob(filespec) and not os.path.exists(filespec):
            yield None, filespec, PathNotFoundException(filespec)
            continue
        for filename in iglob(filespec):
            if zip_password is not None:
                # Each file is expected to be a zip archive:
                # The zip password must be bytes, not unicode/str:
                if not isinstance(zip_password, bytes):
                    zip_password = bytes(zip_password, encoding='utf8')
                # print('Opening zip archive %s with provided password' % filename)
                # print('zip password: %r' % zip_password)
                # print(type(zip_password))
                z = zipfile.ZipFile(filename, 'r')
                #print 'Looking for file(s) matching "%s"' % zip_fname
                for subfilename in ziglob(z, zip_fname):
                    #print 'Opening file in zip archive:', filename
                    try:
                        data = z.read(subfilename, zip_password)
                        yield filename, subfilename, data
                    except Exception as e:
                        yield filename, subfilename, e
                z.close()
            else:
                # normal file
                # do not read the file content, just yield the filename
                yield None, filename, None
                #print 'Opening file', filename
                #data = open(filename, 'rb').read()
                #yield None, filename, data


def is_glob(filespec):
    """ determine if given file specification is a single file name or a glob

    python's glob and fnmatch can only interpret ?, *, [list], and [ra-nge],
    (and combinations: hex_*_[A-Fabcdef0-9]).
    The special chars *?[-] can only be escaped using []
    --> file_name is not a glob
    --> file?name is a glob
    --> file* is a glob
    --> file[-._]name is a glob
    --> file[?]name is not a glob (matches literal "file?name")
    --> file[*]name is not a glob (matches literal "file*name")
    --> file[-]name is not a glob (matches literal "file-name")
    --> file-name is not a glob

    Also, obviously incorrect globs are treated as non-globs
    --> file[name is not a glob (matches literal "file[name")
    --> file]-[name is treated as a glob
        (it is not a valid glob but detecting errors like this requires
         sophisticated regular expression matching)

    Python's glob also works with globs in directory-part of path
    --> dir-part of path is analyzed just like filename-part
    --> thirdparty/*/xglob.py is a (valid) glob
    
    TODO: create a correct regexp to test for validity of ranges
    """

    # remove escaped special chars
    cleaned = filespec.replace('[*]', '').replace('[?]', '') \
                      .replace('[[]', '').replace('[]]', '').replace('[-]', '')

    # check if special chars remain
    return '*' in cleaned or '?' in cleaned or \
          ('[' in cleaned and ']' in cleaned)
