import logging
import json


class JsonFormatter(logging.Formatter):
    """
    Format every message to be logged as a JSON object
    """
    _is_first_line = True

    def format(self, record):
        """
        Since we don't buffer messages, we always prepend messages with a comma to make
        the output JSON-compatible. The only exception is when printing the first line,
        so we need to keep track of it.
        """
        json_dict = dict(msg=record.msg, level=record.levelname)
        formatted_message = '    ' + json.dumps(json_dict)

        if self._is_first_line:
            self._is_first_line = False
            return formatted_message

        return ', ' + formatted_message
