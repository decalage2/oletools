""" Test log_helpers """

import sys
from tests.util.log_helper import log_helper_test_imported
from oletools.util.log_helper import log_helper

DEBUG_MESSAGE = 'main: debug log'
INFO_MESSAGE = 'main: info log'
WARNING_MESSAGE = 'main: warning log'
ERROR_MESSAGE = 'main: error log'
CRITICAL_MESSAGE = 'main: critical log'

logger = log_helper.get_or_create_silent_logger('test_main')


def init_logging_and_log(args):
    """
    Try to cover possible logging scenarios. For each scenario covered, here's the expected args and outcome:
    - Log without enabling: ['<level>']
        * logging when being imported - should never print
    - Log as JSON without enabling: ['as-json', '<level>']
        * logging as JSON when being imported - should never print
    - Enable and log: ['enable', '<level>']
        * logging when being run as script - should log messages
    - Enable and log as JSON: ['as-json', 'enable', '<level>']
        * logging as JSON when being run as script - should log messages as JSON
    - Enable, log as JSON and throw: ['enable', 'as-json', 'throw', '<level>']
        * should produce JSON-compatible output, even after an unhandled exception
    """

    # the level should always be the last argument passed
    level = args[-1]
    use_json = 'as-json' in args
    throw = 'throw' in args

    if 'enable' in args:
        log_helper.enable_logging(use_json, level, stream=sys.stdout)

    _log()

    if throw:
        raise Exception('An exception occurred before ending the logging')

    log_helper.end_logging()


def _log():
    logger.debug(DEBUG_MESSAGE)
    logger.info(INFO_MESSAGE)
    logger.warning(WARNING_MESSAGE)
    logger.error(ERROR_MESSAGE)
    logger.critical(CRITICAL_MESSAGE)
    log_helper_test_imported.log()


if __name__ == '__main__':
    init_logging_and_log(sys.argv[1:])
