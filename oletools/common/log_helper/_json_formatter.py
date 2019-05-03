import logging
import json


class JsonFormatter(logging.Formatter):
    """
    Format every message to be logged as a JSON object
    """
    _is_first_line = True

    def __init__(self, other_logger_has_first_line=False):
        if other_logger_has_first_line:
            self._is_first_line = False

    def format(self, record):
        """
        Since we don't buffer messages, we always prepend messages with a comma to make
        the output JSON-compatible. The only exception is when printing the first line,
        so we need to keep track of it.

        The resulting text is just a json dump of the :py:class:`logging.LogRecord`
        object that is received as input, so no %-formatting or similar is done. Raw
        unformatted message and formatting arguments are contained in fields `msg` and
        `args` of the output.

        Arg `record` has a `type` field when created by `OletoolLoggerAdapter`. If not
        (e.g. captured warnings or output from third-party libraries), we add one.
        """
        json_dict = dict(msg=record.msg.replace('\n', ' '), level=record.levelname)
        try:
            json_dict['type'] = record.type
        except AttributeError:
            if record.name == 'py.warnings':   # this is the name of the logger
                json_dict['type'] = 'warning'
            else:
                json_dict['type'] = 'msg'
        formatted_message = '    ' + json.dumps(json_dict)

        if self._is_first_line:
            self._is_first_line = False
            return formatted_message

        return ', ' + formatted_message
