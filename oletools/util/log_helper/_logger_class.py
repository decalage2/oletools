import logging
from . import _root_logger_wrapper


class OletoolsLogger(logging.Logger):
    """
    Default class for all loggers returned by the logging module.
    """
    def __init__(self, name, level=logging.NOTSET):
        super(self.__class__, self).__init__(name, level)

    def log_at_current_level(self, message):
        """
        Logs the message using the current level.
        This is useful for messages that should always appear,
        such as banners.
        """

        level = _root_logger_wrapper.get_root_logger_level() \
            if _root_logger_wrapper.is_logging_initialized() \
            else self.level

        self.log(level, message)
