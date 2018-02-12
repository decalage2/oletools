import logging


class NullHandler(logging.Handler):
    """
    Log Handler without output
    """

    def emit(self, record):
        pass
