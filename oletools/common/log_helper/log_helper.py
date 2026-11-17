"""
log_helper.py

General logging helpers

Use as follows:

    # at the start of your file:
    # import logging   <-- replace this with next line
    from oletools.common.log_helper import log_helper

    logger = log_helper.get_or_create_silent_logger("module_name")
    def enable_logging():
        '''Enable logging in this module; for use by importing scripts'''
        logger.setLevel(log_helper.NOTSET)
        imported_oletool_module.enable_logging()
        other_imported_oletool_module.enable_logging()

    # ... your code; use logger instead of logging ...

    def main():
        log_helper.enable_logging(level=...)   # instead of logging.basicConfig
        # ... your main code ...
        log_helper.end_logging()

.. codeauthor:: Intra2net AG <info@intra2net>, Philippe Lagadec
"""

# === LICENSE =================================================================

# oletools is copyright (c) 2012-2021, Philippe Lagadec (http://www.decalage.info)
# All rights reserved.
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
#  * Redistributions of source code must retain the above copyright notice,
#    this list of conditions and the following disclaimer.
#  * Redistributions in binary form must reproduce the above copyright notice,
#    this list of conditions and the following disclaimer in the documentation
#    and/or other materials provided with the distribution.
#
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
# AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
# IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
# ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE
# LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
# CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
# SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
# INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
# CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
# ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
# POSSIBILITY OF SUCH DAMAGE.

# -----------------------------------------------------------------------------
# CHANGELOG:
# 2017-12-07 v0.01 CH: - first version
# 2018-02-05 v0.02 SA: - fixed log level selection and reformatted code
# 2018-02-06 v0.03 SA: - refactored code to deal with NullHandlers
# 2018-02-07 v0.04 SA: - fixed control of handlers propagation
# 2018-04-23 v0.05 SA: - refactored the whole logger to use an OOP approach
# 2021-05-17 v0.60 PL: - added default values for enable_logging parameters

# -----------------------------------------------------------------------------
# TODO:


from __future__ import print_function
from ._json_formatter import JsonFormatter
from ._logger_adapter import OletoolsLoggerAdapter
from . import _root_logger_wrapper
from ..io_encoding import ensure_stdout_handles_unicode
import logging
import sys


LOG_LEVELS = {
    'debug': logging.DEBUG,
    'info': logging.INFO,
    'warning': logging.WARNING,
    'error': logging.ERROR,
    'critical': logging.CRITICAL
}

#: provide this constant to modules, so they do not have to import
#: :py:mod:`logging` for themselves just for this one constant.
NOTSET = logging.NOTSET

DEFAULT_LOGGER_NAME = 'oletools'
DEFAULT_MESSAGE_FORMAT = '%(levelname)-8s %(message)s'


class LogHelper:
    """
    Single helper class that creates and remembers loggers.
    """

    #: for convenience: here again (see also :py:data:`log_helper.NOTSET`)
    NOTSET = logging.NOTSET

    def __init__(self):
        self._all_names = set()  # set so we do not have duplicates
        self._use_json = False
        self._is_enabled = False
        self._target_stream = None

    def get_or_create_silent_logger(self, name=DEFAULT_LOGGER_NAME, level=logging.CRITICAL + 1):
        """
        Get a logger or create one if it doesn't exist, setting a NullHandler
        as the handler (to avoid printing to the console).
        By default we also use a higher logging level so every message will
        be ignored.
        This will prevent oletools from logging unnecessarily when being imported
        from external tools.
        """
        return self._get_or_create_logger(name, level, logging.NullHandler())

    def enable_logging(self, use_json=False, level='warning', log_format=DEFAULT_MESSAGE_FORMAT, stream=None,
                       other_logger_has_first_line=False):
        """
        This function initializes the root logger and enables logging.
        We set the level of the root logger to the one passed by calling logging.basicConfig.
        We also set the level of every logger we created to 0 (logging.NOTSET), meaning that
        the level of the root logger will be used to tell if messages should be logged.
        Additionally, since our loggers use the NullHandler, they won't log anything themselves,
        but due to having propagation enabled they will pass messages to the root logger,
        which in turn will log to the stream set in this function.
        Since the root logger is the one doing the work, when using JSON we set its formatter
        so that every message logged is JSON-compatible.

        If other code also creates json output, all items should be pre-pended
        with a comma like the `JsonFormatter` does. Except the first; use param
        `other_logger_has_first_line` to clarify whether our logger or the
        other code will produce the first json item.
        """
        if self._is_enabled:
            raise ValueError('re-enabling logging. Not sure whether that is ok...')

        if stream is None:
            self.target_stream = sys.stdout
        else:
            self.target_stream = stream

        if self.target_stream == sys.stdout:
            ensure_stdout_handles_unicode()

        log_level = LOG_LEVELS[level]
        logging.basicConfig(level=log_level, format=log_format,
                            stream=self.target_stream)
        self._is_enabled = True

        self._use_json = use_json
        sys.excepthook = self._get_except_hook(sys.excepthook)

        # make sure warnings do not mess up our output
        logging.captureWarnings(True)
        warn_logger = self.get_or_create_silent_logger('py.warnings')
        warn_logger.set_warnings_logger()

        # since there could be loggers already created we go through all of them
        # and set their levels to 0 so they will use the root logger's level
        for name in self._all_names:
            logger = self.get_or_create_silent_logger(name)
            self._set_logger_level(logger, logging.NOTSET)

        # add a JSON formatter to the root logger, which will be used by every logger
        if self._use_json:
            _root_logger_wrapper.set_formatter(JsonFormatter(other_logger_has_first_line))
            print('[', file=self.target_stream)

    def end_logging(self):
        """
        Must be called at the end of the main function if the caller wants
        json-compatible output
        """
        if not self._is_enabled:
            return
        self._is_enabled = False

        # end logging
        self._all_names = set()
        logging.captureWarnings(False)
        logging.shutdown()

        # end json list
        if self._use_json:
            print(']', file=self.target_stream)
        self._use_json = False

    def _get_except_hook(self, old_hook):
        """
        Global hook for exceptions so we can always end logging.
        We wrap any hook currently set to avoid overwriting global hooks set by oletools.
        Note that this is only called by enable_logging, which in turn is called by
        the main() function in oletools' scripts. When scripts are being imported this
        code won't execute and won't affect global hooks.
        """
        def hook(exctype, value, traceback):
            self.end_logging()
            old_hook(exctype, value, traceback)

        return hook

    def _get_or_create_logger(self, name, level, handler=None):
        """
        Get or create a new logger. This newly created logger will have the
        handler and level that was passed, but if it already exists it's not changed.
        We also wrap the logger in an adapter so we can easily extend its functionality.
        """

        # logging.getLogger creates a logger if it doesn't exist,
        # so we need to check before calling it
        if handler and not self._log_exists(name):
            logger = logging.getLogger(name)
            logger.addHandler(handler)
            self._set_logger_level(logger, level)
        else:
            logger = logging.getLogger(name)

        # Keep track of every logger we created so we can easily change
        # their levels whenever needed
        self._all_names.add(name)

        adapted_logger = OletoolsLoggerAdapter(logger, None)
        adapted_logger.set_json_enabled_function(lambda: self._use_json)

        return adapted_logger

    @staticmethod
    def _set_logger_level(logger, level):
        """
        If the logging is already initialized, we set the level of our logger
        to 0, meaning that it will reuse the level of the root logger.
        That means that if the root logger level changes, we will keep using
        its level and not logging unnecessarily.
        """

        # if this log was wrapped, unwrap it to set the level
        if isinstance(logger, OletoolsLoggerAdapter):
            logger = logger.logger

        if _root_logger_wrapper.is_logging_initialized():
            logger.setLevel(logging.NOTSET)
        else:
            logger.setLevel(level)

    @staticmethod
    def _log_exists(name):
        """
        We check the log manager instead of our global _all_names variable
        since the logger could have been created outside of the helper
        """
        return name in logging.Logger.manager.loggerDict
