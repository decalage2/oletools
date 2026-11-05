"""
Tests for detect_autoexec keyword classification in olevba.

Covers the interaction-only control events that were previously described as
"Runs when the file is opened ..." (see issue #891): detection must not change,
only the description, so that downstream consumers do not read a user-clicked
button as an auto-executing entry point.
"""

import unittest

from oletools.olevba import detect_autoexec


# Descriptions that assert the code runs when the document is opened. Used to
# check that interaction-only events are not reported with one of these.
ON_OPEN_MARKERS = ('file is opened', 'document is opened', 'Workbook is opened')

CLICK_DESCRIPTION = (
    'Runs when the user interacts with an embedded ActiveX control '
    '(e.g. clicks a button)'
)


class TestAutoexecInteractionEvents(unittest.TestCase):
    """detect_autoexec must describe user-interaction events accurately"""

    def test_click_event_is_not_described_as_on_open(self):
        """A _Click handler is described as interaction, not as on-open."""
        code = 'Private Sub CommandButton1_Click()\nEnd Sub\n'
        results = detect_autoexec(code)

        self.assertEqual(len(results), 1, msg='expected exactly one match: %r' % (results, ))
        keyword, description = results[0]
        self.assertEqual(keyword, 'CommandButton1_Click')
        self.assertEqual(description, CLICK_DESCRIPTION)
        for marker in ON_OPEN_MARKERS:
            self.assertNotIn(marker, description)

    def test_click_event_is_still_detected(self):
        """Splitting the description must not drop the detection from #499."""
        code = 'Sub CommandButton1_Click()\nEnd Sub'
        keywords = [keyword for keyword, _desc in detect_autoexec(code)]
        self.assertIn('CommandButton1_Click', keywords)

    def test_gotfocus_event_keeps_the_on_open_description(self):
        """Other control events still use the on-open description."""
        code = 'Sub TextBox1_GotFocus()\nEnd Sub'
        results = detect_autoexec(code)
        self.assertEqual(len(results), 1, msg='expected exactly one match: %r' % (results, ))
        _keyword, description = results[0]
        self.assertEqual(
            description, 'Runs when the file is opened and ActiveX objects trigger events'
        )

    def test_workbook_open_keeps_the_on_open_description(self):
        """Plain workbook-open handlers are unaffected by this change."""
        code = 'Private Sub Workbook_Open()\nEnd Sub'
        results = detect_autoexec(code)
        self.assertEqual(len(results), 1, msg='expected exactly one match: %r' % (results, ))
        _keyword, description = results[0]
        self.assertEqual(description, 'Runs when the Excel Workbook is opened')

    def test_click_and_gotfocus_are_reported_separately(self):
        """Both events are reported, each with its own description."""
        code = 'Sub CommandButton1_Click()\nEnd Sub\nSub TextBox1_GotFocus()\nEnd Sub'
        results = dict(detect_autoexec(code))
        self.assertEqual(results.get('CommandButton1_Click'), CLICK_DESCRIPTION)
        self.assertEqual(
            results.get('TextBox1_GotFocus'),
            'Runs when the file is opened and ActiveX objects trigger events',
        )


if __name__ == '__main__':
    unittest.main()
