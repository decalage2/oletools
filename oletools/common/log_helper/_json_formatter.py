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

        We assume that all input comes from the OletoolsLoggerAdapter which
        ensures that there is a `type` field in the record. Otherwise will have
        to add a try-except around the access to `record.type`.
        """
        json_dict = dict(msg=record.msg.replace('\n', ' '), level=record.levelname)
        json_dict['type'] = record.type
        formatted_message = '    ' + json.dumps(json_dict)

        if self._is_first_line:
            self._is_first_line = False
            return formatted_message

        return ', ' + formatted_message
