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

MAIN_LOG_MESSAGES = [
    log_helper_test_main.DEBUG_MESSAGE,
    log_helper_test_main.INFO_MESSAGE,
    log_helper_test_main.WARNING_MESSAGE,
    log_helper_test_main.ERROR_MESSAGE,
    log_helper_test_main.CRITICAL_MESSAGE
]


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
            self.assertEquals(len(json_data), len(messages))

            for i in range(len(messages)):
                self.assertEquals(messages[i], json_data[i]['msg'])
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

        self.assertEquals(child.returncode == 0, should_succeed)

        return output.strip()


# just in case somebody calls this file as a script
if __name__ == '__main__':
    unittest.main()
