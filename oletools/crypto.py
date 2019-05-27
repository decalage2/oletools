#!/usr/bin/env python
"""
crypto.py

Module to be used by other scripts and modules in oletools, that provides
information on encryption in OLE files.

Uses :py:mod:`msoffcrypto-tool` to decrypt if it is available. Otherwise
decryption will fail with an ImportError.

Encryption/Write-Protection can be realized in many different ways. They range
from setting a single flag in an otherwise unprotected file to embedding a
regular file (e.g.  xlsx) in an EncryptedStream inside an OLE file. That means
that (1) that lots of bad things are accesible even if no encryption password
is known, and (2) even basic attributes like the file type can change by
decryption. Therefore I suggest the following general routine to deal with
potentially encrypted files::

    def script_main_function(input_file, passwords, crypto_nesting=0, args):
        '''Wrapper around main function to deal with encrypted files.'''
        initial_stuff(input_file, args)
        result = None
        try:
            result = do_your_thing_assuming_no_encryption(input_file)
            if not crypto.is_encrypted(input_file):
                return result
        except Exception:
            if not crypto.is_encrypted(input_file):
                raise
        # we reach this point only if file is encrypted
        # check if this is an encrypted file in an encrypted file in an ...
        if crypto_nesting >= crypto.MAX_NESTING_DEPTH:
            raise crypto.MaxCryptoNestingReached(crypto_nesting, filename)
        decrypted_file = None
        try:
            decrypted_file = crypto.decrypt(input_file, passwords)
            if decrypted_file is None:
                raise crypto.WrongEncryptionPassword(input_file)
            # might still be encrypted, so call this again recursively
            result = script_main_function(decrypted_file, passwords,
                                          crypto_nesting+1, args)
        except Exception:
            raise
        finally:     # clean up
            try:     # (maybe file was not yet created)
                os.unlink(decrypted_file)
            except Exception:
                pass

(Realized e.g. in :py:mod:`oletools.msodde`).
That means that caller code needs another wrapper around its main function. I
did try it another way first (a transparent on-demand unencrypt) but for the
above reasons I believe this is the better way. Also, non-top-level-code can
just assume that it works on unencrypted data and fail with an exception if
encrypted data makes its work impossible. No need to check `if is_encrypted()`
at the start of functions.

.. seealso:: [MS-OFFCRYPTO]
.. seealso:: https://github.com/nolze/msoffcrypto-tool

crypto is part of the python-oletools package:
http://www.decalage.info/python/oletools
"""

# === LICENSE =================================================================

# crypto is copyright (c) 2014-2021 Philippe Lagadec (http://www.decalage.info)
# All rights reserved.
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
#  * Redistributions of source code must retain the above copyright notice,
#    this list of conditions and the following disclaimer.
#  * Redistributions in binary form must reproduce the above copyright notice,
#    this list of conditions and the following disclaimer in the documentation
#    and/or other materials provided with the distribution.
#
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
# AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
# IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
# ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE
# LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
# CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
# SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
# INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
# CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
# ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
# POSSIBILITY OF SUCH DAMAGE.

# -----------------------------------------------------------------------------
# CHANGELOG:
# 2019-02-14 v0.01 CH: - first version with encryption check from oleid
# 2019-04-01 v0.54 PL: - fixed bug in is_encrypted_ole
# 2019-05-23       PL: - added DEFAULT_PASSWORDS list
# 2021-05-22 v0.60 PL: - added PowerPoint transparent password
#                        '/01Hannes Ruescher/01' (issue #627)
# 2019-05-24       CH: - use log_helper

__version__ = '0.60'

import sys
import struct
import os
from os.path import splitext, isfile
from tempfile import mkstemp
import zipfile

from olefile import OleFileIO

try:
    import msoffcrypto
except ImportError:
    msoffcrypto = None

# IMPORTANT: it should be possible to run oletools directly as scripts
# in any directory without installing them with pip or setup.py.
# In that case, relative imports are NOT usable.
# And to enable Python 2+3 compatibility, we need to use absolute imports,
# so we add the oletools parent folder to sys.path (absolute+normalized path):
_thismodule_dir = os.path.normpath(os.path.abspath(os.path.dirname(__file__)))
_parent_dir = os.path.normpath(os.path.join(_thismodule_dir, '..'))
if _parent_dir not in sys.path:
    sys.path.insert(0, _parent_dir)

from oletools.common.errors import CryptoErrorBase, WrongEncryptionPassword, \
    UnsupportedEncryptionError, MaxCryptoNestingReached, CryptoLibNotImported
from oletools.common.log_helper import log_helper


#: if there is an encrypted file embedded in an encrypted file,
#: how deep down do we go
MAX_NESTING_DEPTH = 10

# === LOGGING =================================================================

# a global logger object used for debugging:
log = log_helper.get_or_create_silent_logger('crypto')


def enable_logging():
    """
    Enable logging for this module (disabled by default).

    For use by third-party libraries that import `crypto` as module.

    This will set the module-specific logger level to NOTSET, which
    means the main application controls the actual logging level.
    """
    log.setLevel(log_helper.NOTSET)


def is_encrypted(some_file):
    """
    Determine whether document contains encrypted content.

    This should return False for documents that are just write-protected or
    signed or finalized. It should return True if ANY content of the file is
    encrypted and can therefore not be analyzed by other oletools modules
    without given a password.

    Exception: there are way to write-protect an office document by embedding
    it as encrypted stream with hard-coded standard password into an otherwise
    empty OLE file. From an office user point of view, this is no encryption,
    but regarding file structure this is encryption, so we return `True` for
    these.

    This should not raise exceptions needlessly.

    This implementation is rather simple: it returns True if the file contains
    streams with typical encryption names (c.f. [MS-OFFCRYPTO]). It does not
    test whether these streams actually contain data or whether the ole file
    structure contains the necessary references to these. It also checks the
    "well-known property" PIDSI_DOC_SECURITY if the SummaryInformation stream
    is accessible (c.f. [MS-OLEPS] 2.25.1)

    :param some_file: File name or an opened OleFileIO
    :type some_file: :py:class:`olefile.OleFileIO` or `str`
    :returns: True if (and only if) the file contains encrypted content
    """
    # ask msoffcrypto if possible
    if check_msoffcrypto():
        log.debug('Checking for encryption using msoffcrypto')
        file_handle = None
        file_pos = None
        try:
            if isinstance(some_file, OleFileIO):
                # TODO: hacky, replace once msoffcrypto-tools accepts OleFileIO
                file_handle = some_file.fp
                file_pos = file_handle.tell()
                file_handle.seek(0)
            else:
                file_handle = open(some_file, 'rb')

            return msoffcrypto.OfficeFile(file_handle).is_encrypted()

        except Exception as exc:
            # TODO: this triggers unnecessary warnings for non OLE files
            log.info('msoffcrypto failed to parse file or determine '
                        'whether it is encrypted: {}'
                        .format(exc))
            # TODO: here we are ignoring some exceptions that should be raised, for example
            #       "unknown file format" for Excel 5.0/95 files

        finally:
            try:
                if file_pos is not None:   # input was OleFileIO
                    file_handle.seek(file_pos)
                else:                      # input was file name
                    file_handle.close()
            except Exception as exc:
                log.warning('Ignoring error during clean up: {}'.format(exc))

    # if that failed, try ourselves with older and less accurate code
    try:
        if isinstance(some_file, OleFileIO):
            return _is_encrypted_ole(some_file)
        if zipfile.is_zipfile(some_file):
            return _is_encrypted_zip(some_file)
        # otherwise assume it is the name of an ole file
        with OleFileIO(some_file) as ole:
            return _is_encrypted_ole(ole)
    except Exception as exc:
        # TODO: this triggers unnecessary warnings for non OLE files
        log.info('Failed to check {} for encryption ({}); assume it is not '
                    'encrypted.'.format(some_file, exc))

    return False


def _is_encrypted_zip(filename):
    """Specialization of :py:func:`is_encrypted` for zip-based files."""
    log.debug('Checking for encryption in zip file')
    # TODO: distinguish OpenXML from normal zip files
    # try to decrypt a few bytes from first entry
    with zipfile.ZipFile(filename, 'r') as zipper:
        first_entry = zipper.infolist()[0]
        try:
            with zipper.open(first_entry, 'r') as reader:
                reader.read(min(16, first_entry.file_size))
            return False
        except RuntimeError as rt_err:
            return 'crypt' in str(rt_err)


def _is_encrypted_ole(ole):
    """Specialization of :py:func:`is_encrypted` for ole files."""
    log.debug('Checking for encryption in OLE file')
    # check well known property for password protection
    # (this field may be missing for Powerpoint2000, for example)
    # TODO: check whether password protection always implies encryption. Could
    #       write-protection or signing with password trigger this as well?
    if ole.exists("\x05SummaryInformation"):
        suminfo_data = ole.getproperties("\x05SummaryInformation")
        if 0x13 in suminfo_data and (suminfo_data[0x13] & 1):
            return True

    # check a few stream names
    # TODO: check whether these actually contain data and whether other
    # necessary properties exist / are set
    if ole.exists('EncryptionInfo'):
        log.debug('found stream EncryptionInfo')
        return True
    # or an encrypted ppt file
    if ole.exists('EncryptedSummary') and \
            not ole.exists('SummaryInformation'):
        return True

    # Word-specific old encryption:
    if ole.exists('WordDocument'):
        # check for Word-specific encryption flag:
        stream = None
        try:
            stream = ole.openstream(["WordDocument"])
            # pass header 10 bytes
            stream.read(10)
            # read flag structure:
            temp16 = struct.unpack("H", stream.read(2))[0]
            f_encrypted = (temp16 & 0x0100) >> 8
            if f_encrypted:
                return True
        finally:
            if stream is not None:
                stream.close()

    # no indication of encryption
    return False


#: one way to achieve "write protection" in Excel files is to encrypt the file
#: using this password
# ref: https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/6b4a08cb-195a-442e-b31c-7c94624a8c29#Appendix_A_25
# ref: https://twitter.com/BouncyHat/status/1308897568773083138
EXCEL_TRANSPARENT_PASSWORD = 'VelvetSweatshop'

# PowerPoint password which is transparent for the user:
# ref: https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/57fc02f0-c1de-4fc6-908f-d146104662f5
# ref: https://twitter.com/BouncyHat/status/1308897932389896192
POWERPOINT_TRANSPARENT_PASSWORD = '/01Hannes Ruescher/01'

#: list of common passwords to be tried by default, used by malware
DEFAULT_PASSWORDS = [EXCEL_TRANSPARENT_PASSWORD, POWERPOINT_TRANSPARENT_PASSWORD,
                     '123', '1234', '12345', '123456', '4321']


def _check_msoffcrypto():
    """Raise a :py:class:`CryptoLibNotImported` if msoffcrypto not imported."""
    if msoffcrypto is None:
        raise CryptoLibNotImported()


def check_msoffcrypto():
    """Return `True` iff :py:mod:`msoffcrypto` could be imported."""
    return msoffcrypto is not None


def decrypt(filename, passwords=None, **temp_file_args):
    """
    Try to decrypt an encrypted file

    This function tries to decrypt the given file using a given set of
    passwords. If no password is given, tries the standard password for write
    protection. Creates a file with decrypted data whose file name is returned.
    If the decryption fails, None is returned.

    :param str filename: path to an ole file on disc
    :param passwords: list/set/tuple/... of passwords or a single password or
                      None
    :type passwords: iterable or str or None
    :param temp_file_args: arguments for :py:func:`tempfile.mkstemp` e.g.,
                           `dirname` or `prefix`. `suffix` will default to
                           suffix of input `filename`, `prefix` defaults to
                           `oletools-decrypt-`; `text` will be ignored
    :returns: name of the decrypted temporary file (type str) or `None`
    :raises: :py:class:`ImportError` if :py:mod:`msoffcrypto-tools` not found
    :raises: :py:class:`ValueError` if the given file is not encrypted
    """
    _check_msoffcrypto()

    # normalize password so we always have a list/tuple
    if isinstance(passwords, str):
        passwords = (passwords, )
    elif not passwords:
        passwords = DEFAULT_PASSWORDS

    # check temp file args
    if 'prefix' not in temp_file_args:
        temp_file_args['prefix'] = 'oletools-decrypt-'
    if 'suffix' not in temp_file_args:
        temp_file_args['suffix'] = splitext(filename)[1]
    temp_file_args['text'] = False

    decrypt_file = None
    with open(filename, 'rb') as reader:
        try:
            crypto_file = msoffcrypto.OfficeFile(reader)
        except Exception as exc:   # e.g. ppt, not yet supported by msoffcrypto
            if 'Unrecognized file format' in str(exc):
                log.debug('Caught exception', exc_info=True)

                # raise different exception without stack trace of original exc
                if sys.version_info.major == 2:
                    raise UnsupportedEncryptionError(filename)
                else:
                    # this is a syntax error in python 2, so wrap it in exec()
                    exec('raise UnsupportedEncryptionError(filename) from None')
            else:
                raise
        if not crypto_file.is_encrypted():
            raise ValueError('Given input file {} is not encrypted!'
                             .format(filename))

        for password in passwords:
            log.debug('Trying to decrypt with password {!r}'.format(password))
            write_descriptor = None
            write_handle = None
            decrypt_file = None
            try:
                crypto_file.load_key(password=password)

                # create temp file
                write_descriptor, decrypt_file = mkstemp(**temp_file_args)
                write_handle = os.fdopen(write_descriptor, 'wb')
                write_descriptor = None      # is now handled via write_handle
                crypto_file.decrypt(write_handle)

                # decryption was successfull; clean up and return
                write_handle.close()
                write_handle = None
                break
            except Exception:
                log.debug('Failed to decrypt', exc_info=True)

                # error-clean up: close everything and del temp file
                if write_handle:
                    write_handle.close()
                elif write_descriptor:
                    os.close(write_descriptor)
                if decrypt_file and isfile(decrypt_file):
                    os.unlink(decrypt_file)
                decrypt_file = None
    # if we reach this, all passwords were tried without success
    log.debug('All passwords failed')
    return decrypt_file
