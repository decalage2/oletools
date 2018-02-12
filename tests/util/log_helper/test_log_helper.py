""" Test the log helper

This tests the generic log helper.
Check if it handles imported modules correctly
and that the default silent logger won't log when nothing is enabled
"""

import unittest
import sys
import json
import re
from tests.util.log_helper import log_helper_test_main
from tests.util.log_helper import log_helper_test_imported
from os.path import dirname, join
from subprocess import check_output, STDOUT, CalledProcessError

ROOT_DIRECTORY = dirname(dirname(dirname(__file__)))
TEST_FILE = join(dirname(__file__), 'log_helper_test_main.py')
PYTHON_EXECUTABLE = sys.executable
REGEX = re.compile('<#(.*)(:?#>|Traceback)', re.MULTILINE | re.DOTALL)

MAIN_LOG_MESSAGES = [
    log_helper_test_main.DEBUG_MESSAGE,
    log_helper_test_main.INFO_MESSAGE,
    log_helper_test_main.WARNING_MESSAGE,
    log_helper_test_main.ERROR_MESSAGE,
    log_helper_test_main.CRITICAL_MESSAGE
]


class TestLogHelper(unittest.TestCase):
    def test_default_logging_to_stderr(self):
        """
        Basic test for simple logging
        """
        output = self._run_test(['default'])
        self.assertIn(log_helper_test_main.WARNING_MESSAGE, output)

    def test_logging_silently(self):
        """
        Test that nothing will be logged when logging is not enabled
        and we are using a silent logger (uses the NullHandler)
        """
        output = self._run_test(['silent'])
        self.assertTrue(len(output) == 0)

    def test_setting_level_in_main_module(self):
        """
        Make sure that the level set in the main module is kept when
        logging from imported modules.
        """
        output = self._run_test(['debug'])

        expected_messages = MAIN_LOG_MESSAGES + [
            log_helper_test_imported.DEBUG_MESSAGE,
            log_helper_test_imported.INFO_MESSAGE,
            log_helper_test_imported.WARNING_MESSAGE,
            log_helper_test_imported.ERROR_MESSAGE,
            log_helper_test_imported.CRITICAL_MESSAGE
        ]

        for msg in expected_messages:
            self.assertIn(msg, output)

    def test_logging_at_current_level(self):
        """
        Test that logging at current level will always print a message
        """
        output = self._run_test(['current_level'])
        self.assertIn(log_helper_test_main.DEBUG_MESSAGE, output)

    def test_logging_as_json(self):
        """
        Basic test for json logging
        """
        output = self._run_test(['critical', '-j'])

        try:
            json_data = json.loads(output)
            self._assert_json_messages(json_data, [
                log_helper_test_main.CRITICAL_MESSAGE,
                log_helper_test_imported.CRITICAL_MESSAGE
            ])
        except ValueError:
            self.fail('Invalid json:\n' + output)
        self.assertNotEqual(len(json_data), 0, msg='Output was empty')

    def test_logging_dictionary_as_json(self):
        """
        Test support for passing a dictionary to the logger
        and have it logged as JSON
        """
        output = self._run_test(['dictionary'])

        try:
            json_data = json.loads(output)
            self._assert_json_messages(json_data, [
                log_helper_test_main.DEBUG_MESSAGE
            ])
        except ValueError:
            self.fail('Invalid json:\n' + output)
        self.assertNotEqual(len(json_data), 0, msg='Output was empty')

    def test_json_correct_on_exceptions(self):
        """
        Test that even on unhandled exceptions our JSON is always correct
        """
        output = self._run_test(['critical', 'throw', '-j'], True)

        try:
            json_data = json.loads(output)
            self._assert_json_messages(json_data, [
                log_helper_test_main.CRITICAL_MESSAGE,
                log_helper_test_imported.CRITICAL_MESSAGE
            ])
        except ValueError:
            self.fail('Invalid json:\n' + output)
        self.assertNotEqual(len(json_data), 0, msg='Output was empty')

    def _assert_json_messages(self, json_data, messages):
        self.assertEquals(len(json_data), len(messages))

        for i in range(len(messages)):
            self.assertEquals(messages[i], json_data[i]['msg'])

    @staticmethod
    def _run_test(args, ignore_exceptions=False):
        """
        Use subprocess to better simulate the real scenario and avoid
        logging conflicts when running multiple tests (since logging depends on singletons,
        we might get errors or false positives between sequential tests runs)
        """
        try:
            output = check_output(
                [PYTHON_EXECUTABLE, TEST_FILE] + args,
                shell=False,
                cwd=ROOT_DIRECTORY,
                stderr=STDOUT,
                universal_newlines=True
            )

            if not isinstance(output, str):
                output = output.decode('utf-8')
        except CalledProcessError as ex:
            if ignore_exceptions:
                output = ex.output
            else:
                # we want tests to fail if an exception occur
                raise ex

        return REGEX.search(output).group(1).strip()


# just in case somebody calls this file as a script
if __name__ == '__main__':
    unittest.main()
