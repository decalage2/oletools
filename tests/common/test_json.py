"""
Test that all --json output is always valid json.

Since this test takes rather long, it is not included in regular unittest runs.
To enable it, set environment variable OLETOOLS_TEST_JSON to value "1"
"""

import os
from os.path import relpath
import json
import unittest

from tests.test_utils import DATA_BASE_DIR, call_and_capture
from tests.test_utils.testdata_reader import loop_and_extract


@unittest.skipIf('OLETOOLS_TEST_JSON' not in os.environ or os.environ['OLETOOLS_TEST_JSON'] != '1',
                 'Test takes pretty long, do not include in regular test runs')
class TestJson(unittest.TestCase):
    """Test that all --json output is always valid json."""

    def test_all(self):
        """Check that olevba, msodde and oleobj produce valid json for ALL samples."""
        for sample_path in loop_and_extract():
            if sample_path.startswith(DATA_BASE_DIR):
                print(f'TestJson: checking sample {relpath(sample_path, DATA_BASE_DIR)}')
            else:
                print(f'TestJson: checking sample {sample_path}')
            output, _ = call_and_capture('oleobj', ['--json', '--nodump', sample_path],
                                         accept_nonzero_exit=True)
            json.loads(output)

            output, _ = call_and_capture('olevba', ['--json', sample_path],
                                         accept_nonzero_exit=True)
            json.loads(output)

            output, _ = call_and_capture('msodde', ['--json', sample_path],
                                         accept_nonzero_exit=True)
            json.loads(output)


# just in case somebody calls this file as a script
if __name__ == '__main__':
    unittest.main()
