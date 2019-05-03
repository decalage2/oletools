import logging
from . import _root_logger_wrapper


class OletoolsLoggerAdapter(logging.LoggerAdapter):
    """
    Adapter class for all loggers returned by the logging module.
    """
    _json_enabled = None
    _is_warn_logger = False   # this is always False

    def print_str(self, message, **kwargs):
        """
        This function replaces normal print() calls so we can format them as JSON
        when needed or just print them right away otherwise.
        """
        if self._json_enabled and self._json_enabled():
            # Messages from this function should always be printed,
            # so when using JSON we log using the same level that set.
            # Additional information in kwargs is added to LogRecord
            self.log(_root_logger_wrapper.level(), message, extra=kwargs)
        else:
            print(message)

    def log(self, lvl, msg, *args, **kwargs):
        """
        Run :py:meth:`process` on kwargs, then forward to actual logger.

        This is based on the logging cookbox, section "Using LoggerAdapter to
        impart contextual information".
        """
        msg, kwargs = self.process(msg, kwargs)
        self.logger.log(lvl, msg, *args, **kwargs)

    def process(self, msg, kwargs):
        """
        Ensure `kwargs['extra']['type']` exists, init with given arg `type`.

        The `type` field will be added to the :py:class:`logging.LogRecord` and
        is used by the :py:class:`JsonFormatter`.
        """
        if 'extra' not in kwargs:
            kwargs['extra'] = {}
        if 'type' in kwargs:
            kwargs['extra']['type'] = kwargs['type']
            del kwargs['type']    # downstream loggers cannot deal with this
        if 'type' not in kwargs['extra']:
            if self._is_warn_logger:
                kwargs['extra']['type'] = 'warning'    # this will add field
            else:
                kwargs['extra']['type'] = 'msg'        # 'type' to LogRecord
        return msg, kwargs

    def set_json_enabled_function(self, json_enabled):
        """
        Set a function to be called to check whether JSON output is enabled.
        """
        self._json_enabled = json_enabled

    def set_warnings_logger(self):
        """Make this the logger for warnings"""
        # create a object attribute that shadows the class attribute which is
        # always False
        self._is_warn_logger = True

    def level(self):
        """Return current level of logger."""
        return self.logger.level

    def setLevel(self, new_level):
        """Set level of underlying logger. Required only for python < 3.2."""
        return self.logger.setLevel(new_level)
