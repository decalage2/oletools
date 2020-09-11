#!/usr/bin/env python3

"""
Module for testing import of common logging modules by third party modules.

This module behaves like a third party module. It does not use the common
logging and enables logging on its own. But it imports log_helper_test_main.
"""

import sys
import logging

from tests.common.log_helper import log_helper_test_main


def main(args):
    """
    Main function, called when running file as script

    see module doc for more info
    """
    logging.basicConfig(level=logging.INFO)
    if 'enable' in args:
        log_helper_test_main.enable_logging()

    logging.debug('Should not show.')
    logging.info('Start message from 3rd party importer')

    log_helper_test_main.do_log()

    logging.debug('Returning 0, but you will never see that ... .')
    logging.info('End message from 3rd party importer')
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
