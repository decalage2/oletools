""" class OutputCapture to test what scripts print to stdout """

from __future__ import print_function
import sys


# python 2/3 version conflict:
if sys.version_info.major <= 2:
    from StringIO import StringIO
else:
    from io import StringIO

class OutputCapture:
    """ context manager that captures stdout

    use as follows::

        with OutputCapture() as capturer:
            run_my_script(some_args)

        # either test line-by-line ...
        for line in capturer:
            some_test(line)
        # ...or test all output in one go
        some_test(capturer.get_data())

    """

    def __init__(self):
        self.buffer = StringIO()
        self.orig_stdout = None
        self.data = None

    def __enter__(self):
        # replace sys.stdout with own buffer.
        self.orig_stdout = sys.stdout
        sys.stdout = self.buffer
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        sys.stdout = self.orig_stdout    # re-set to original
        self.data = self.buffer.getvalue()
        self.buffer.close()              # close buffer
        self.buffer = None

        if exc_type:   # there has been an error
            print('Got error during output capture!')
            print('Print captured output and re-raise:')
            for line in self.data.splitlines():
                print(line.rstrip())  # print output before re-raising

    def get_data(self):
        """ retrieve all the captured data """
        if self.buffer is not None:
            return self.buffer.getvalue()
        elif self.data is not None:
            return self.data
        else:   # should not be possible
            raise RuntimeError('programming error or someone messed with data!')

    def __iter__(self):
        for line in self.get_data().splitlines():
            yield line
