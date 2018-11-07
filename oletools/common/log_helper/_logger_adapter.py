import logging
from . import _root_logger_wrapper


class OletoolsLoggerAdapter(logging.LoggerAdapter):
    """
    Adapter class for all loggers returned by the logging module.
    """
    _json_enabled = None

    def print_str(self, message):
        """
        This function replaces normal print() calls so we can format them as JSON
        when needed or just print them right away otherwise.
        """
        if self._json_enabled and self._json_enabled():
            # Messages from this function should always be printed,
            # so when using JSON we log using the same level that set
            self.log(_root_logger_wrapper.level(), message)
        else:
            print(message)

    def set_json_enabled_function(self, json_enabled):
        """
        Set a function to be called to check whether JSON output is enabled.
        """
        self._json_enabled = json_enabled

    def level(self):
        return self.logger.level
