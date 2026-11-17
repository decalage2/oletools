import logging
import json


class JsonFormatter(logging.Formatter):
    """
    Format every message to be logged as a JSON object

    Uses the standard :py:class:`logging.Formatter` with standard arguments
    to do the actual formatting, could save and use a user-supplied formatter
    instead.
    """
    _is_first_line = True

    def __init__(self, other_logger_has_first_line=False):
        if other_logger_has_first_line:
            self._is_first_line = False
        self.msg_formatter = logging.Formatter()    # could adjust this

    def format(self, record):
        """
        Since we don't buffer messages, we always prepend messages with a comma to make
        the output JSON-compatible. The only exception is when printing the first line,
        so we need to keep track of it.

        The actual conversion from :py:class:`logging.LogRecord` to a text message
        (i.e. %-formatting, adding exception information, etc.) is delegated to the
        standard :py:class:`logging.Formatter.

        The dumped json structure contains fields `msg` with the formatted message,
        `level` with the log-level of the message and `type`, which is created by
        :py:class:`oletools.common.log_helper.OletoolsLoggerAdapter` or added here
        (for input from e.g. captured warnings, third-party libraries)
        """
        json_dict = dict(msg='', level='', type='')
        try:
            msg = self.msg_formatter.format(record)
            json_dict['msg'] = msg.replace('\n', ' ')
            json_dict['level'] = record.levelname
            json_dict['type'] = record.type
        except AttributeError:     # most probably: record has no "type" field
            if record.name == 'py.warnings':   # this is from python's warning-capture logger
                json_dict['type'] = 'warning'
            else:
                json_dict['type'] = 'msg'      # message of unknown origin
        except Exception as exc:
            try:
                json_dict['msg'] = "Ignore {0} when formatting '{1}': {2}".format(type(exc), record.msg, exc)
            except Exception as exc2:
                json_dict['msg'] = 'Caught {0} in logging'.format(str(exc2))
            json_dict['type'] = 'log-warning'
            json_dict['level'] = 'warning'

        formatted_message = '    ' + json.dumps(json_dict)

        if self._is_first_line:
            self._is_first_line = False
            return formatted_message

        return ', ' + formatted_message
