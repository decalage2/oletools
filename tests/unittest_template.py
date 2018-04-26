""" Test my new feature

Some more info if you want

Should work with python2 and python3!
"""

import unittest

# if you need data from oletools/test-data/DIR/, uncomment these lines:
## Directory with test data, independent of current working directory
#from tests.test_utils import DATA_BASE_DIR


class TestMyFeature(unittest.TestCase):
    """ Tests my cool new feature """

    def test_this(self):
        """ check that this works """
        pass   # your code here

    def test_that(self):
        """ check that that also works """
        pass   # your code here

    def helper_function(self, filename):
        """ to be called from other test functions to avoid copy-and-paste

        this is not called by unittest directly, only from your functions """
        pass   # your code here
        # e.g.: msodde.main(join(DATA_DIR, filename))


# just in case somebody calls this file as a script
if __name__ == '__main__':
    unittest.main()
