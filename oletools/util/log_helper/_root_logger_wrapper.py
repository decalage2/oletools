import logging


def is_logging_initialized():
    """
    We use the same strategy as the logging module
    when checking if the logging was initialized -
    look for handlers in the root logger
    """
    return len(logging.root.handlers) > 0


def get_root_logger_level():
    return logging.root.level
