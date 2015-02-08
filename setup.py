"""Installs oletools using distutils

Run:
    python setup.py install

to install this package.

(setup script partly borrowed from cherrypy)
"""

#--- CHANGELOG ----------------------------------------------------------------

# 2014-08-27 v0.06: - added doc subfolder
# 2015-01-05 v0.07: - added xglob, prettytable
# 2015-02-08 v0.08: - added DridexUrlDecoder

#--- TODO ---------------------------------------------------------------------

# - deploy scripts?
# - version for python 3 using 2to3

#--- IMPORTS ------------------------------------------------------------------

##try:
##    from setuptools import setup
##except ImportError:
from distutils.core import setup

#from distutils.command.install import INSTALL_SCHEMES

import sys, os, fnmatch


#--- METADATA -----------------------------------------------------------------

name         = "oletools"
version      = '0.08'
desc         = "Python tools to analyze security characteristics of MS OLE2 files (also called Structured Storage, Compound File Binary Format or Compound Document File Format), such as Microsoft Office documents, for Malware Analysis and Incident Response."
long_desc    = open('oletools/README.rst').read()
author       ="Philippe Lagadec"
author_email ="decalage at laposte dot net"
url          ="http://www.decalage.info/python/oletools"
license      ="BSD"
download_url ="https://bitbucket.org/decalage/oletools/downloads"

classifiers=[
    "Development Status :: 3 - Alpha",
    "Intended Audience :: Information Technology",
    "Intended Audience :: Science/Research",
    "Intended Audience :: System Administrators",
    "License :: OSI Approved :: BSD License",
    "Operating System :: OS Independent",
    "Programming Language :: Python",
    "Programming Language :: Python :: 2",
#    "Programming Language :: Python :: 3",
    "Topic :: Security",
]

#--- PACKAGES -----------------------------------------------------------------

packages=[
    "oletools",
    'oletools.thirdparty',
    'oletools.thirdparty.olefile',
    'oletools.thirdparty.easygui',
    'oletools.thirdparty.xxxswf',
    'oletools.thirdparty.prettytable',
    'oletools.thirdparty.xglob',
    'oletools.thirdparty.DridexUrlDecoder',
]
##setupdir = '.'
##package_dir={'': setupdir}

#--- PACKAGE DATA -------------------------------------------------------------

## Often, additional files need to be installed into a package. These files are
## often data that?s closely related to the package?s implementation, or text
## files containing documentation that might be of interest to programmers using
## the package. These files are called package data.
##
## Package data can be added to packages using the package_data keyword argument
## to the setup() function. The value must be a mapping from package name to a
## list of relative path names that should be copied into the package. The paths
## are interpreted as relative to the directory containing the package
## (information from the package_dir mapping is used if appropriate); that is,
## the files are expected to be part of the package in the source directories.
## They may contain glob patterns as well.
##
## The path names may contain directory portions; any necessary directories will
## be created in the installation.


# the following functions are used to dynamically include package data without
# listing every file here:

def riglob(top, prefix='', pattern='*'):
    """
    recursive iterator glob
    - top: path to start searching from
    - prefix: path to use instead of top when generating file path (in order to
      choose the root of relative paths)
    - pattern: wilcards to select files (same syntax as fnmatch)

    Yields each file found in top and subdirectories, matching pattern
    """
    #print 'top=%s prefix=%s pat=%s' % (top, prefix, pattern)
    dirs = []
    for path in os.listdir(top):
        p = os.path.join(top, path)
        if os.path.isdir(p):
            dirs.append(path)
        elif os.path.isfile(p):
            #print ' - file:', path
            if fnmatch.fnmatch(path, pattern):
                yield os.path.join(prefix, path)
    #print ' dirs =', dirs
    for d in dirs:
        dtop = os.path.join(top, d)
        dprefix = os.path.join(prefix, d)
        #print 'dtop=%s dprefix=%s' % (dtop, dprefix)
        for p in riglob(dtop, dprefix, pattern):
            yield p

def rglob(top, prefix='', pattern='*'):
    """
    recursive glob
    Same as riglob, but returns a list.
    """
    return list(riglob(top, prefix, pattern))




package_data={
    'oletools': [
        'README.rst',
        'README.html',
        'LICENSE.txt',
        ]
        # doc folder: md, html, png
        + rglob('oletools/doc', 'doc', '*.html')
        + rglob('oletools/doc', 'doc', '*.md')
        + rglob('oletools/doc', 'doc', '*.png'),

    'oletools.thirdparty.olefile': [
        'README.txt',
        'LICENSE.txt',
        ],
    'oletools.thirdparty.xglob': [
        'LICENSE.txt',
        ],
    'oletools.thirdparty.easygui': [
        'LICENSE.txt',
        ],
    'oletools.thirdparty.xxxswf': [
        'LICENSE.txt',
        ],
    'oletools.thirdparty.prettytable': [
        'CHANGELOG', 'COPYING', 'README'
        ],
    'oletools.thirdparty.DridexUrlDecoder': [
        'LICENSE.txt',
        ],
    }


#--- data files ---------------------------------------------------------------

# not used for now.

## The data_files option can be used to specify additional files needed by the
## module distribution: configuration files, message catalogs, data files,
## anything which doesn?t fit in the previous categories.
##
## data_files specifies a sequence of (directory, files) pairs in the following way:
##
## setup(...,
##       data_files=[('bitmaps', ['bm/b1.gif', 'bm/b2.gif']),
##                   ('config', ['cfg/data.cfg']),
##                   ('/etc/init.d', ['init-script'])]
##      )
##
## Note that you can specify the directory names where the data files will be
## installed, but you cannot rename the data files themselves.
##
## Each (directory, files) pair in the sequence specifies the installation
## directory and the files to install there. If directory is a relative path,
## it is interpreted relative to the installation prefix (Python?s sys.prefix for
## pure-Python packages, sys.exec_prefix for packages that contain extension
## modules). Each file name in files is interpreted relative to the setup.py
## script at the top of the package source distribution. No directory information
## from files is used to determine the final location of the installed file;
## only the name of the file is used.
##
## You can specify the data_files options as a simple sequence of files without
## specifying a target directory, but this is not recommended, and the install
## command will print a warning in this case. To install data files directly in
## the target directory, an empty string should be given as the directory.



##data_files=[
##    ('balbuzard', [
##        'balbuzard/README.txt',
##                  ]),
##]

##if sys.version_info >= (3, 0):
##    required_python_version = '3.0'
##    setupdir = 'py3'
##else:
##    required_python_version = '2.3'
##    setupdir = 'py2'

##data_files = [(install_dir, ['%s/%s' % (setupdir, f) for f in files])
##              for install_dir, files in data_files]


##def fix_data_files(data_files):
##    """
##    bdist_wininst seems to have a bug about where it installs data files.
##    I found a fix the django team used to work around the problem at
##    http://code.djangoproject.com/changeset/8313 .  This function
##    re-implements that solution.
##    Also see http://mail.python.org/pipermail/distutils-sig/2004-August/004134.html
##    for more info.
##    """
##    def fix_dest_path(path):
##        return '\\PURELIB\\%(path)s' % vars()
##
##    if not 'bdist_wininst' in sys.argv: return
##
##    data_files[:] = [
##        (fix_dest_path(path), files)
##        for path, files in data_files]
##fix_data_files(data_files)


#--- SCRIPTS ------------------------------------------------------------------

# not used for now.

#scripts = ["%s/cherrypy/cherryd" % setupdir]


#=== MAIN =====================================================================

def main():
    if sys.version >= '3.0':
        s = "Sorry, %s %s requires Python 2.x."
        print(s % (name, version))
        sys.exit(1)
##    if sys.version < required_python_version:
##        s = "I'm sorry, but %s %s requires Python %s or later."
##        print(s % (name, version, required_python_version))
##        sys.exit(1)
##    # set default location for "data_files" to
##    # platform specific "site-packages" location
##    for scheme in list(INSTALL_SCHEMES.values()):
##        scheme['data'] = scheme['purelib']

    dist = setup(
        name=name,
        version=version,
        description=desc,
        long_description=long_desc,
        classifiers=classifiers,
        author=author,
        author_email=author_email,
        url=url,
        license=license,
##        package_dir=package_dir,
        packages=packages,
        package_data = package_data,
        download_url=download_url,
#        data_files=data_files,
#        scripts=scripts,
    )


if __name__ == "__main__":
    main()

