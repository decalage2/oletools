import logging


def is_logging_initialized():
    """
    We use the same strategy as the logging module when checking if
    the logging was initialized - look for handlers in the root logger
    """
    return len(logging.root.handlers) > 0


def set_formatter(fmt):
    """
    Set the formatter to be used by every handler of the root logger.
    """
    if not is_logging_initialized():
        return

    for handler in logging.root.handlers:
        handler.setFormatter(fmt)


def level():
    return logging.root.level
