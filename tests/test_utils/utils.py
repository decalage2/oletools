#!/usr/bin/env python3

"""Utils generally useful for unittests."""

import sys
import os
from os.path import dirname, join, abspath
from subprocess import check_output, PIPE, STDOUT, CalledProcessError


# Base dir of project, contains subdirs "tests" and "oletools" and README.md
PROJECT_ROOT = dirname(dirname(dirname(abspath(__file__))))

# Directory with test data, independent of current working directory
DATA_BASE_DIR = join(PROJECT_ROOT, 'tests', 'test-data')

# Directory with source code
SOURCE_BASE_DIR = join(PROJECT_ROOT, 'oletools')


def call_and_capture(module, args=None, accept_nonzero_exit=False,
                     exclude_stderr=False):
    """
    Run module as script, capturing and returning output and return code.

    This is the best way to capture a module's stdout and stderr; trying to
    modify sys.stdout/sys.stderr to StringIO-Buffers frequently causes trouble.

    Only drawback sofar: stdout and stderr are merged into one (which is
    what users see on their shell as well). When testing for json-compatible
    output you should `exclude_stderr` to `False` since logging ignores stderr,
    so unforseen warnings (e.g. issued by pypy) would mess up your json.

    :param str module: name of module to test, e.g. `olevba`
    :param args: arguments for module's main function
    :param bool fail_nonzero: Raise error if command returns non-0 return code
    :param bool exclude_stderr: Exclude output to `sys.stderr` from output
                                (e.g. if parsing output through json)
    :returns: ret_code, output
    :rtype: int, str
    """
    # create a PYTHONPATH environment var to prefer our current code
    env = os.environ.copy()
    try:
        env['PYTHONPATH'] = SOURCE_BASE_DIR + os.pathsep + \
                            os.environ['PYTHONPATH']
    except KeyError:
        env['PYTHONPATH'] = SOURCE_BASE_DIR

    # hack: in python2 output encoding (sys.stdout.encoding) was None
    # although sys.getdefaultencoding() and sys.getfilesystemencoding were ok
    # TODO: maybe can remove this once branch
    #       "encoding-for-non-unicode-environments" is merged
    if 'PYTHONIOENCODING' not in env:
        env['PYTHONIOENCODING'] = 'utf8'

    # ensure args is a tuple
    my_args = tuple(args) if args else ()

    ret_code = -1
    try:
        output = check_output((sys.executable, '-m', module) + my_args,
                              universal_newlines=True, env=env,
                              stderr=PIPE if exclude_stderr else STDOUT)
        ret_code = 0

    except CalledProcessError as err:
        if accept_nonzero_exit:
            ret_code = err.returncode
            output = err.output
        else:
            print(err.output)
            raise

    return output, ret_code
