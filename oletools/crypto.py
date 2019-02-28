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

    def script_main_function(input_file, args):
        '''Wrapper around main function to deal with encrypted files.'''
        initial_stuff(input_file, args)
        result = None
        try:
            result = do_your_thing_assuming_no_encryption(input_file)
            if not crypto_is_encrypted(input_file):
                return result
        except Exception:
            if not crypto_is_encrypted(input_file):
                raise
        decrypted_file = None
        try:
            decrypted_file = crypto.decrypt(input_file)
        except Exception:
            raise
        finally:     # clean up
            try:     # (maybe file was not yet created)
                os.unlink(decrypted_file)
            except Exception:
                pass

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

# crypto is copyright (c) 2014-2019 Philippe Lagadec (http://www.decalage.info)
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

__version__ = '0.01'

import struct
import os
from os.path import splitext, isfile
from tempfile import mkstemp

try:
    import msoffcrypto
except ImportError:
    msoffcrypto = None


def is_encrypted(olefile):
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

    :param olefile: An opened OleFileIO or a filename to such a file
    :type olefile: :py:class:`olefile.OleFileIO` or `str`
    :returns: True if (and only if) the file contains encrypted content
    """
    if isinstance(olefile, str):
        ole = olefile.OleFileIO(olefile)
    else:
        ole = olefile   # assume it is an olefile.OleFileIO

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
    elif ole.exists('EncryptionInfo'):
        return True
    # or an encrypted ppt file
    elif ole.exists('EncryptedSummary') and \
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
        except Exception:
            raise
        finally:
            if stream is not None:
                stream.close()

    # no indication of encryption
    return False


#: one way to achieve "write protection" in office files is to encrypt the file
#: using this password
WRITE_PROTECT_ENCRYPTION_PASSWORD = 'VelvetSweatshop'


def _check_msoffcrypto():
    """raise an :py:class:`ImportError` if :py:data:`msoffcrypto` is `None`."""
    if msoffcrypto is None:
        raise ImportError('msoffcrypto-tools could not be imported')


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
                           suffix of input `filename`; `text` will be ignored
    :returns: name of the decrypted temporary file.
    :raises: :py:class:`ImportError` if :py:mod:`msoffcrypto-tools` not found
    :raises: :py:class:`ValueError` if the given file is not encrypted
    """
    _check_msoffcrypto()

    if passwords is None:
        passwords = (WRITE_PROTECT_ENCRYPTION_PASSWORD, )
    elif isinstance(passwords, str):
        passwords = (passwords, )

    decrypt_file = None
    with open(filename, 'rb') as reader:
        crypto_file = msoffcrypto.OfficeFile(reader)
        if not crypto_file.is_encrypted():
            raise ValueError('Given input file {} is not encrypted!'
                             .format(filename))

        for password in passwords:
            try:
                crypto_file.load_key(password=password)
            except Exception:
                continue     # password verification failed, try next

            # create temp file
            if 'suffix' not in temp_file_args:
                temp_file_args['suffix'] = splitext(filename)[1]
            temp_file_args['text'] = False

            write_descriptor = None
            write_handle = None
            try:
                write_descriptor, decrypt_file = mkstemp(**temp_file_args)
                write_handle = os.fdopen(write_descriptor, 'wb')
                write_descriptor = None      # is now handled via write_handle
                crypto_file.decrypt(write_handle)

                # decryption was successfull; clean up and return
                write_handle.close()
                write_handle = None
                break
            except Exception:
                # error: clean up: close everything and del file ignoring errors;
                # then re-raise original exception
                if write_handle:
                    write_handle.close()
                elif write_descriptor:
                    os.close(write_descriptor)
                if decrypt_file and isfile(decrypt_file):
                    os.unlink(decrypt_file)
                decrypt_file = None
                raise
    # if we reach this, all passwords were tried without success
    return decrypt_file
