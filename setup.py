#!/usr/bin/env python
"""
Installs oletools using distutils

Run:
    python setup.py install

to install this package.

(setup script partly borrowed from cherrypy)
"""

#--- CHANGELOG ----------------------------------------------------------------

# 2014-08-27 v0.06 PL: - added doc subfolder
# 2015-01-05 v0.07 PL: - added xglob, prettytable
# 2015-02-08 v0.08 PL: - added DridexUrlDecoder
# 2015-03-23 v0.09 PL: - updated description and classifiers, added shebang line
# 2015-06-16 v0.10 PL: - added pyparsing
# 2016-02-08 v0.42 PL: - added colorclass, tablestream
# 2016-07-19 v0.50 PL: - create CLI scripts using entry points (by 2*yo)
# 2016-07-29       PL: - use setuptools if available
# 2016-09-05       PL: - added more entry points
# 2017-01-18 v0.51 PL: - added package zipfile27 (issue #121)
# 2017-10-18 v0.52 PL: - added msodde
# 2018-03-19 v0.52.3      PL: - added install_requires, removed thirdparty.pyparsing
# 2018-09-11 v0.54 PL: - olefile is now a dependency
# 2018-09-15       PL: - easygui is now a dependency
# 2018-09-22       PL: - colorclass is now a dependency
# 2018-10-27       PL: - fixed issue #359 (bug when importing log_helper)
# 2019-02-26       CH: - add optional dependency msoffcrypto for decryption
# 2019-05-22       PL: - 'msoffcrypto-tool' is now a required dependency
# 2019-05-23 v0.55 PL: - added pcodedmp as dependency
# 2019-09-24       PL: - removed oletools.thirdparty.DridexUrlDecoder
# 2019-11-10       PL: - changed pyparsing from 2.2.0 to 2.1.0 for issue #481

#--- TODO ---------------------------------------------------------------------


#--- IMPORTS ------------------------------------------------------------------

try:
    from setuptools import setup
except ImportError:
    from distutils.core import setup

#from distutils.command.install import INSTALL_SCHEMES

import os, fnmatch


#--- METADATA -----------------------------------------------------------------

name         = "oletools"
version      = '0.56dev6'
desc         = "Python tools to analyze security characteristics of MS Office and OLE files (also called Structured Storage, Compound File Binary Format or Compound Document File Format), for Malware Analysis and Incident Response #DFIR"
long_desc    = open('oletools/README.rst').read()
author       = "Philippe Lagadec"
author_email = "nospam@decalage.info"
url          = "http://www.decalage.info/python/oletools"
license      = "BSD"
download_url = "https://github.com/decalage2/oletools/releases"

# see https://pypi.org/pypi?%3Aaction=list_classifiers
classifiers=[
    "Development Status :: 4 - Beta",
    "Intended Audience :: Developers",
    "Intended Audience :: Information Technology",
    "Intended Audience :: Science/Research",
    "Intended Audience :: System Administrators",
    "License :: OSI Approved :: BSD License",
    "Natural Language :: English",
    "Operating System :: OS Independent",
    "Programming Language :: Python",
    "Programming Language :: Python :: 2",
    "Programming Language :: Python :: 2.7",
    "Programming Language :: Python :: 3",
    "Programming Language :: Python :: 3.4",
    "Programming Language :: Python :: 3.5",
    "Programming Language :: Python :: 3.6",
    "Programming Language :: Python :: 3.7",
    "Programming Language :: Python :: 3.8",
    "Topic :: Security",
    "Topic :: Software Development :: Libraries :: Python Modules",
]

#--- PACKAGES -----------------------------------------------------------------

packages=[
    "oletools",
    "oletools.common",
    "oletools.common.log_helper",
    'oletools.thirdparty',
    'oletools.thirdparty.xxxswf',
    'oletools.thirdparty.prettytable',
    'oletools.thirdparty.xglob',
    'oletools.thirdparty.tablestream',
    'oletools.thirdparty.oledump',
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

    'oletools.thirdparty.xglob': [
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
    # 'oletools.thirdparty.tablestream': [
    #     'LICENSE', 'README',
    #     ],
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


# --- SCRIPTS ------------------------------------------------------------------

# Entry points to create convenient scripts automatically

entry_points = {
    'console_scripts': [
        'ezhexviewer=oletools.ezhexviewer:main',
        'mraptor=oletools.mraptor:main',
        'mraptor3=oletools.mraptor3:main',
        'olebrowse=oletools.olebrowse:main',
        'oledir=oletools.oledir:main',
        'oleid=oletools.oleid:main',
        'olemap=oletools.olemap:main',
        'olemeta=oletools.olemeta:main',
        'oletimes=oletools.oletimes:main',
        'olevba=oletools.olevba:main',
        'olevba3=oletools.olevba3:main',
        'pyxswf=oletools.pyxswf:main',
        'rtfobj=oletools.rtfobj:main',
        'oleobj=oletools.oleobj:main',
        'msodde=oletools.msodde:main',
        'olefile=olefile.olefile:main',
    ],
}

# scripts=['oletools/olevba.py', 'oletools/mraptor.py']


# === MAIN =====================================================================

def main():
    # TODO: warning about Python 2.6
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
        # package_dir=package_dir,
        packages=packages,
        package_data = package_data,
        download_url=download_url,
        # data_files=data_files,
        entry_points=entry_points,
        test_suite="tests",
        # scripts=scripts,
        install_requires=[
            "pyparsing>=2.1.0,<3",  # changed from 2.2.0 to 2.1.0 for issue #481
            "olefile>=0.46",
            "easygui",
            'colorclass',
            'msoffcrypto-tool',
            'pcodedmp>=1.2.5',
        ],
    )


if __name__ == "__main__":
    main()

