import logging
import json


class JsonFormatter(logging.Formatter):
    """
    Format every message to be logged as a JSON object
    """
    def __init__(self, fmt=None, datefmt=None):
        super(JsonFormatter, self).__init__(fmt, datefmt)
        self._is_first_line = True

    def format(self, record):
        """
        We accept messages that are either dictionaries or not.
        When we have dictionaries we can just serialize it as JSON right away.
        """
        trailing_comma = ','

        if self._is_first_line:
            trailing_comma = ''
            self._is_first_line = False

        json_dict = record.msg \
            if isinstance(record.msg, dict) \
            else dict(msg=record.msg)
        json_dict['level'] = record.levelname

        return trailing_comma + '    ' + json.dumps(json_dict)
