""" Test the log helper

This tests the generic log helper.
Check if it handles imported modules correctly
and that the default silent logger won't log when nothing is enabled
"""

import unittest
import sys
import json
import subprocess
from tests.common.log_helper import log_helper_test_main
from tests.common.log_helper import log_helper_test_imported
from os.path import dirname, join, relpath, abspath

from tests.test_utils import PROJECT_ROOT

# this is the common base of "tests" and "oletools" dirs
TEST_FILE = relpath(join(dirname(abspath(__file__)), 'log_helper_test_main.py'),
                    PROJECT_ROOT)
PYTHON_EXECUTABLE = sys.executable


class TestLogHelper(unittest.TestCase):
    def test_it_doesnt_log_when_not_enabled(self):
        output = self._run_test(['debug'])
        self.assertTrue(len(output) == 0)

    def test_it_doesnt_log_json_when_not_enabled(self):
        output = self._run_test(['as-json', 'debug'])
        self.assertTrue(len(output) == 0)

    def test_logs_when_enabled(self):
        output = self._run_test(['enable', 'warning'])

        expected_messages = [
            log_helper_test_main.WARNING_MESSAGE,
            log_helper_test_main.ERROR_MESSAGE,
            log_helper_test_main.CRITICAL_MESSAGE,
            log_helper_test_imported.WARNING_MESSAGE,
            log_helper_test_imported.ERROR_MESSAGE,
            log_helper_test_imported.CRITICAL_MESSAGE
        ]

        for msg in expected_messages:
            self.assertIn(msg, output)

    def test_logs_json_when_enabled(self):
        output = self._run_test(['enable', 'as-json', 'critical'])

        self._assert_json_messages(output, [
            log_helper_test_main.CRITICAL_MESSAGE,
            log_helper_test_imported.CRITICAL_MESSAGE
        ])

    def test_logs_type_ignored(self):
        """Run test script with logging enabled at info level. Want no type."""
        output = self._run_test(['enable', 'info'])

        expect = '\n'.join([
            'INFO     ' + log_helper_test_main.INFO_MESSAGE,
            'WARNING  ' + log_helper_test_main.WARNING_MESSAGE,
            'ERROR    ' + log_helper_test_main.ERROR_MESSAGE,
            'CRITICAL ' + log_helper_test_main.CRITICAL_MESSAGE,
            'INFO     ' + log_helper_test_main.RESULT_MESSAGE,
            'INFO     ' + log_helper_test_imported.INFO_MESSAGE,
            'WARNING  ' + log_helper_test_imported.WARNING_MESSAGE,
            'ERROR    ' + log_helper_test_imported.ERROR_MESSAGE,
            'CRITICAL ' + log_helper_test_imported.CRITICAL_MESSAGE,
            'INFO     ' + log_helper_test_imported.RESULT_MESSAGE,
        ])
        self.assertEqual(output, expect)

    def test_logs_type_in_json(self):
        """Check type field is contained in json log."""
        output = self._run_test(['enable', 'as-json', 'info'])

        # convert to json preserving order of output
        jout = json.loads(output)

        jexpect = [
            dict(type='msg', level='INFO',
                 msg=log_helper_test_main.INFO_MESSAGE),
            dict(type='msg', level='WARNING',
                 msg=log_helper_test_main.WARNING_MESSAGE),
            dict(type='msg', level='ERROR',
                 msg=log_helper_test_main.ERROR_MESSAGE),
            dict(type='msg', level='CRITICAL',
                 msg=log_helper_test_main.CRITICAL_MESSAGE),
            # this is the important entry (has a different "type" field):
            dict(type=log_helper_test_main.RESULT_TYPE, level='INFO',
                 msg=log_helper_test_main.RESULT_MESSAGE),
            dict(type='msg', level='INFO',
                 msg=log_helper_test_imported.INFO_MESSAGE),
            dict(type='msg', level='WARNING',
                 msg=log_helper_test_imported.WARNING_MESSAGE),
            dict(type='msg', level='ERROR',
                 msg=log_helper_test_imported.ERROR_MESSAGE),
            dict(type='msg', level='CRITICAL',
                 msg=log_helper_test_imported.CRITICAL_MESSAGE),
            # ... and this:
            dict(type=log_helper_test_imported.RESULT_TYPE, level='INFO',
                 msg=log_helper_test_imported.RESULT_MESSAGE),
        ]
        self.assertEqual(jout, jexpect)

    def test_percent_autoformat(self):
        """Test that auto-formatting of log strings with `%` works."""
        output = self._run_test(['enable', '%-autoformat', 'info'])
        self.assertIn('The answer is 47.', output)

    def test_json_correct_on_exceptions(self):
        """
        Test that even on unhandled exceptions our JSON is always correct
        """
        output = self._run_test(['enable', 'as-json', 'throw', 'critical'], False)
        self._assert_json_messages(output, [
            log_helper_test_main.CRITICAL_MESSAGE,
            log_helper_test_imported.CRITICAL_MESSAGE
        ])

    def _assert_json_messages(self, output, messages):
        try:
            json_data = json.loads(output)
            self.assertEqual(len(json_data), len(messages))

            for i in range(len(messages)):
                self.assertEqual(messages[i], json_data[i]['msg'])
        except ValueError:
            self.fail('Invalid json:\n' + output)

        self.assertNotEqual(len(json_data), 0, msg='Output was empty')

    def _run_test(self, args, should_succeed=True):
        """
        Use subprocess to better simulate the real scenario and avoid
        logging conflicts when running multiple tests (since logging depends on singletons,
        we might get errors or false positives between sequential tests runs)
        """
        child = subprocess.Popen(
            [PYTHON_EXECUTABLE, TEST_FILE] + args,
            shell=False,
            env={'PYTHONPATH': PROJECT_ROOT},
            universal_newlines=True,
            cwd=PROJECT_ROOT,
            stdin=None,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )
        (output, output_err) = child.communicate()

        if not isinstance(output, str):
            output = output.decode('utf-8')

        self.assertEqual(child.returncode == 0, should_succeed)

        return output.strip()


# just in case somebody calls this file as a script
if __name__ == '__main__':
    unittest.main()
