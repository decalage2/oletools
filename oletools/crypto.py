#!/usr/bin/env python
"""
crypto.py

Module to be used by other scripts and modules in oletools, that provides
information on encryption in OLE files.

.. seealso:: [MS-OFFCRYPTO]

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
