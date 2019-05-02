"""
Dummy file that logs messages, meant to be imported
by the main test file
"""

from oletools.common.log_helper import log_helper
import warnings

DEBUG_MESSAGE = 'imported: debug log'
INFO_MESSAGE = 'imported: info log'
WARNING_MESSAGE = 'imported: warning log'
ERROR_MESSAGE = 'imported: error log'
CRITICAL_MESSAGE = 'imported: critical log'
RESULT_MESSAGE = 'imported: result log'

RESULT_TYPE = 'imported: result'
ACTUAL_WARNING = 'Feature XYZ provided by this module might be deprecated at '\
    'some point in the future ... or not'

logger = log_helper.get_or_create_silent_logger('test_imported')

def enable_logging():
    """Enable logging if imported by third party modules."""
    logger.setLevel(log_helper.NOTSET)


def log():
    logger.debug(DEBUG_MESSAGE)
    logger.info(INFO_MESSAGE)
    logger.warning(WARNING_MESSAGE)
    logger.error(ERROR_MESSAGE)
    logger.critical(CRITICAL_MESSAGE)
    logger.info(RESULT_MESSAGE, type=RESULT_TYPE)


def warn():
    warnings.warn(ACTUAL_WARNING)
