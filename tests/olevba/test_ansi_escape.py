"""
Test detection of ANSI escape codes in VBA code (issue #874).
"""

import unittest

from oletools.olevba import detect_suspicious


class TestAnsiEscapeDetection(unittest.TestCase):
    """Tests that ANSI escape sequences in VBA code are reported as suspicious."""

    VBA_WITH_ANSI = 'Sub AutoOpen()\n' \
        'Dim s As String\n' \
        's = "\x1b[2J\x1b[0;0H"\n' \
        'End Sub\n'

    VBA_WITHOUT_ANSI = 'Sub AutoOpen()\n' \
        'Dim s As String\n' \
        's = "calc"\n' \
        'End Sub\n'

    def test_ansi_escape_detected(self):
        """Plain VBA code containing ANSI escape sequences must be flagged."""
        results = detect_suspicious(self.VBA_WITH_ANSI)
        descriptions = [description for _, description in results]
        self.assertTrue(
            any('ANSI escape' in description for description in descriptions),
            'ANSI escape codes not detected in: %r' % (results,))

    def test_no_raw_esc_in_reported_keyword(self):
        """The reported keyword must never contain a raw ESC character."""
        results = detect_suspicious(self.VBA_WITH_ANSI)
        for keyword, _ in results:
            self.assertNotIn('\x1b', keyword)

    def test_clean_code_not_flagged(self):
        """VBA code without ANSI escapes must not trigger the ANSI warning."""
        results = detect_suspicious(self.VBA_WITHOUT_ANSI)
        descriptions = [description for _, description in results]
        self.assertFalse(
            any('ANSI escape' in description for description in descriptions),
            'False ANSI detection in: %r' % (results,))

    def test_ansi_escape_from_base64_not_reported(self):
        """ANSI escapes found only after Base64 deobfuscation must not be
        reported (same guard as for backspace chars)."""
        results = detect_suspicious(self.VBA_WITH_ANSI, obfuscation='Base64')
        descriptions = [description for _, description in results]
        self.assertFalse(
            any('ANSI escape' in description for description in descriptions),
            'ANSI escape reported from obfuscated content: %r' % (results,))

    def test_ansi_escape_from_hex_not_reported(self):
        """ANSI escapes found only after Hex deobfuscation must not be
        reported (same guard as for backspace chars)."""
        results = detect_suspicious(self.VBA_WITH_ANSI, obfuscation='Hex')
        descriptions = [description for _, description in results]
        self.assertFalse(
            any('ANSI escape' in description for description in descriptions),
            'ANSI escape reported from obfuscated content: %r' % (results,))


if __name__ == '__main__':
    unittest.main()
