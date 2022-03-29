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

# test file we use as "main" module
TEST_FILE = relpath(join(dirname(abspath(__file__)), 'log_helper_test_main.py'),
                    PROJECT_ROOT)

# test file simulating a third party main module that only imports oletools
TEST_FILE_3RD_PARTY = relpath(join(dirname(abspath(__file__)),
                                   'third_party_importer.py'),
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

    def test_import_by_third_party_disabled(self):
        """Test that when imported by third party, logging is still disabled."""
        output = self._run_test([], run_third_party=True).splitlines()
        self.assertEqual(len(output), 2)
        self.assertEqual(output[0],
                         'INFO:root:Start message from 3rd party importer')
        self.assertEqual(output[1],
                         'INFO:root:End message from 3rd party importer')

    def test_import_by_third_party_enabled(self):
        """Test that when imported by third party, logging can be enabled."""
        output = self._run_test(['enable', ], run_third_party=True).splitlines()
        self.assertEqual(len(output), 12)
        self.assertIn('INFO:test_main:main: info log', output)
        self.assertIn('INFO:test_imported:imported: info log', output)

    def _assert_json_messages(self, output, messages):
        try:
            json_data = json.loads(output)
            self.assertEqual(len(json_data), len(messages))

            for i in range(len(messages)):
                self.assertEqual(messages[i], json_data[i]['msg'])
        except ValueError:
            self.fail('Invalid json:\n' + output)

        self.assertNotEqual(len(json_data), 0, msg='Output was empty')

    def _run_test(self, args, should_succeed=True, run_third_party=False):
        """
        Use subprocess to better simulate the real scenario and avoid
        logging conflicts when running multiple tests (since logging depends on singletons,
        we might get errors or false positives between sequential tests runs)

        When arg `run_third_party` is `True`, we do not run the `TEST_FILE` as
        main moduel but the `TEST_FILE_3RD_PARTY` and return contents of
        `stderr` instead of `stdout`.
        """
        all_args = [PYTHON_EXECUTABLE, ]
        if run_third_party:
            all_args.append(TEST_FILE_3RD_PARTY)
        else:
            all_args.append(TEST_FILE)
        all_args.extend(args)
        child = subprocess.Popen(
            all_args,
            shell=False,
            env={'PYTHONPATH': PROJECT_ROOT},
            universal_newlines=True,
            cwd=PROJECT_ROOT,
            stdin=None,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )
        (output, output_err) = child.communicate()

        if False:   # DEBUG
            print()
            for line in output_err.splitlines():
                print('ERR: {}'.format(line.rstrip()))
            for line in output.splitlines():
                print('OUT: {}'.format(line.rstrip()))

        if run_third_party:
            output = output_err

        if not isinstance(output, str):
            output = output.decode('utf-8')

        self.assertEqual(child.returncode == 0, should_succeed)

        return output.strip()


# just in case somebody calls this file as a script
if __name__ == '__main__':
    unittest.main()
