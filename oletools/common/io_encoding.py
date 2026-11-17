#!/usr/bin/env python3

"""
Tool to help with input/output encoding

Helpers to run smoothly in unicode-unfriendly environments like output redirect
or unusual language settings.

In such settings, output to console falls back to ASCII-only. Also open()
suddenly fails to interprete non-ASCII characters.

Therefore, at start of scripts can run :py:meth:`ensure_stdout_handles_unicode`
and when opening text files use :py:meth:`uopen` to replace :py:meth:`open`.

Part of the python-oletools package:
http://www.decalage.info/python/oletools
"""

# === LICENSE =================================================================

# msodde is copyright (c) 2017-2018 Philippe Lagadec (http://www.decalage.info)
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
# 2018-11-04 v0.54 CH: - first version: ensure_stdout_handles_unicode, uopen

# -- IMPORTS ------------------------------------------------------------------
from __future__ import print_function
import sys
import codecs
import os
from locale import getpreferredencoding

PY3 = sys.version_info.major >= 3

if PY3:
    from builtins import open as builtin_open
else:
    from __builtin__ import open as builtin_open    # pylint: disable=import-error

# -- CONSTANTS ----------------------------------------------------------------
#: encoding to use for redirection if no good encoding can be found
FALLBACK_ENCODING_REDIRECT = 'utf8'

#: encoding for reading text from files if preferred encoding is non-unicode
FALLBACK_ENCODING_OPEN = 'utf8'

#: print (pure-ascii) debug output to stdout
DEBUG = False

# the encoding specified in system environment
try:
    PREFERRED_ENCODING = getpreferredencoding()
except Exception as exc:
    if DEBUG:
        print('Exception getting preferred encoding: {}'.format(exc))
    PREFERRED_ENCODING = None


# -- HELPERS =-----------------------------------------------------------------


def ensure_stdout_handles_unicode():
    """
    Ensure that print()ing unicode does not lead to errors.

    When print()ing unicode, python relies on the environment (e.g. in linux on
    the setting of the LANG environment variable) to tell it how to encode
    unicode. That works nicely for modern-day shells where encoding is usually
    UTF-8. But as soon as LANG is unset or just "C", or output is redirected or
    piped, the encoding falls back to 'ASCII', which cannot handle unicode
    characters.

    Based on solutions suggested on stackoverflow (c.f.
    https://stackoverflow.com/q/27347772/4405656 ), wrap stdout in an encoder
    that solves that problem.

    Unfortunately, stderr cannot be handled the same way ( see e.g. https://
    pythonhosted.org/kitchen/unicode-frustrations.html#frustration-5-exceptions
    ), so we still have to hope there is only ascii in error messages
    """
    # do not re-wrap
    if isinstance(sys.stdout, codecs.StreamWriter):
        if DEBUG:
            print('sys.stdout wrapped already')
        return

    # get output stream object
    if PY3:
        output_stream = sys.stdout.buffer
    else:
        output_stream = sys.stdout

    # determine encoding of sys.stdout
    try:
        encoding = sys.stdout.encoding
    except AttributeError:              # variable "encoding" might not exist
        encoding = None
    if DEBUG:
        print('sys.stdout encoding is {}'.format(encoding))

    if isinstance(encoding, str) and encoding.lower().startswith('utf'):
        if DEBUG:
            print('encoding is acceptable')
        return     # everything alright, we are working in a good environment
    elif os.isatty(output_stream.fileno()):   # e.g. C locale
        # Do not output UTF8 since that might be mis-interpreted.
        # Just replace chars that cannot be handled
        print('Encoding for stdout is only {}, will replace other chars to '
              'avoid unicode error'.format(encoding), file=sys.stderr)
        sys.stdout = codecs.getwriter(encoding)(output_stream, errors='replace')
    else:                                  # e.g. redirection, pipe in python2
        new_encoding = PREFERRED_ENCODING
        if DEBUG:
            print('not a tty, try preferred encoding {}'.format(new_encoding))
        if not isinstance(new_encoding, str) \
                or not new_encoding.lower().startswith('utf'):
            new_encoding = FALLBACK_ENCODING_REDIRECT
            if DEBUG:
                print('preferred encoding also unacceptable, fall back to {}'
                      .format(new_encoding))
        print('Encoding for stdout is only {}, will auto-encode text with {} '
              'before output'.format(encoding, new_encoding), file=sys.stderr)
        sys.stdout = codecs.getwriter(new_encoding)(output_stream)


def uopen(filename, mode='r', *args, **kwargs):
    """
    Replacement for builtin open() that reads unicode even in ASCII environment

    In order to read unicode from text, python uses locale.getpreferredencoding
    to translate bytes to str. If the environment only provides ASCII encoding,
    this will fail since most office files contain unicode.

    Therefore, guess a good encoding here if necessary and open file with that.

    :returns: same type as the builtin :py:func:`open`
    """
    # do not interfere if not necessary:
    if 'b' in mode:
        if DEBUG:
            print('Opening binary file, do not interfere')
        return builtin_open(filename, mode, *args, **kwargs)
    if 'encoding' in kwargs:
        if DEBUG:
            print('Opening file with encoding {!r}, do not interfere'
                  .format(kwargs['encoding']))
        return builtin_open(filename, mode, *args, **kwargs)
    if len(args) > 3:    # "encoding" is the 4th arg
        if DEBUG:
            print('Opening file with encoding {!r}, do not interfere'
                  .format(args[3]))
        return builtin_open(filename, mode, *args, **kwargs)

    # determine preferred encoding
    encoding = PREFERRED_ENCODING
    if DEBUG:
        print('preferred encoding is {}'.format(encoding))

    if isinstance(encoding, str) and encoding.lower().startswith('utf'):
        if DEBUG:
            print('encoding is acceptable, open {} regularly'.format(filename))
        return builtin_open(filename, mode, *args, **kwargs)

    # so we want to read text from a file but can probably only deal with ASCII
    # --> use fallback
    if DEBUG:
        print('Opening {} with fallback encoding {}'
              .format(filename, FALLBACK_ENCODING_OPEN))
    if PY3:
        return builtin_open(filename, mode, *args,
                            encoding=FALLBACK_ENCODING_OPEN, **kwargs)
    else:
        handle = builtin_open(filename, mode, *args, **kwargs)
        return codecs.EncodedFile(handle, FALLBACK_ENCODING_OPEN)
