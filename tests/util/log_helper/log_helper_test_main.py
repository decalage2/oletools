""" Test log_helpers """

import sys
import logging
from tests.util.log_helper import log_helper_test_imported
from oletools.util.log_helper import log_helper

DEBUG_MESSAGE = 'main: debug log'
INFO_MESSAGE = 'main: info log'
WARNING_MESSAGE = 'main: warning log'
ERROR_MESSAGE = 'main: error log'
CRITICAL_MESSAGE = 'main: critical log'

logger = log_helper.get_or_create_logger('test_main')


def init_logging_and_log(cmd_line_args=None):
    args = cmd_line_args if cmd_line_args else sys.argv

    if 'silent' in args:
        return _log_silently()
    elif 'dictionary' in args:
        return _log_dictionary(args)
    elif 'current_level' in args:
        _enable_logging()
        return _log_at_current_level()
    elif 'default' in args:
        _enable_logging()
        return _log(logger)

    use_json = '-j' in args
    throw_exception = 'throw' in args

    level = _parse_log_level(args)

    _enable_logging(use_json, level)
    _log(logger)
    log_helper_test_imported.log()

    if throw_exception:
        raise Exception('An exception occurred before ending the logging')

    _end_logging()


def _parse_log_level(args):
    if 'debug' in args:
        return 'debug'
    elif 'info' in args:
        return 'info'
    elif 'warning' in args:
        return 'warning'
    elif 'error' in args:
        return 'error'
    else:
        return 'critical'


def _log_dictionary(args):
    level = _parse_log_level(args)
    _enable_logging(True, level)

    logger.log_at_current_level({
        'msg': DEBUG_MESSAGE
    })
    log_helper.end_logging()


def _enable_logging(use_json=False, level='warning'):
    log_helper.enable_logging(use_json, level, stream=sys.stdout)


def _log_at_current_level():
    logger.log_at_current_level(DEBUG_MESSAGE)


def _log_silently():
    silent_logger = log_helper.get_or_create_silent_logger('test_main_silent', logging.DEBUG - 1)
    _log(silent_logger)


def _log(current_logger):
    current_logger.debug(DEBUG_MESSAGE)
    current_logger.info(INFO_MESSAGE)
    current_logger.warning(WARNING_MESSAGE)
    current_logger.error(ERROR_MESSAGE)
    current_logger.critical(CRITICAL_MESSAGE)


def _end_logging():
    log_helper.end_logging()


if __name__ == '__main__':
    # since we are using subprocess, add delimiters so we can easily extract
    # the output that matters (when debugging tests we get extra output,
    # so we need to ignore it)
    print('<#')
    init_logging_and_log(sys.argv)
    print('#>')
