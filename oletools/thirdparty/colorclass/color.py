"""Color class used by library users."""

from colorclass.core import ColorStr


class Color(ColorStr):
    """Unicode (str in Python3) subclass with ANSI terminal text color support.

    Example syntax: Color('{red}Sample Text{/red}')

    Example without parsing logic: Color('{red}Sample Text{/red}', keep_tags=True)

    For a list of codes, call: colorclass.list_tags()
    """

    @classmethod
    def colorize(cls, color, string, auto=False):
        """Color-code entire string using specified color.

        :param str color: Color of string.
        :param str string: String to colorize.
        :param bool auto: Enable auto-color (dark/light terminal).

        :return: Class instance for colorized string.
        :rtype: Color
        """
        tag = '{0}{1}'.format('auto' if auto else '', color)
        return cls('{%s}%s{/%s}' % (tag, string, tag))

    @classmethod
    def black(cls, string, auto=False):
        """Color-code entire string.

        :param str string: String to colorize.
        :param bool auto: Enable auto-color (dark/light terminal).

        :return: Class instance for colorized string.
        :rtype: Color
        """
        return cls.colorize('black', string, auto=auto)

    @classmethod
    def bgblack(cls, string, auto=False):
        """Color-code entire string.

        :param str string: String to colorize.
        :param bool auto: Enable auto-color (dark/light terminal).

        :return: Class instance for colorized string.
        :rtype: Color
        """
        return cls.colorize('bgblack', string, auto=auto)

    @classmethod
    def red(cls, string, auto=False):
        """Color-code entire string.

        :param str string: String to colorize.
        :param bool auto: Enable auto-color (dark/light terminal).

        :return: Class instance for colorized string.
        :rtype: Color
        """
        return cls.colorize('red', string, auto=auto)

    @classmethod
    def bgred(cls, string, auto=False):
        """Color-code entire string.

        :param str string: String to colorize.
        :param bool auto: Enable auto-color (dark/light terminal).

        :return: Class instance for colorized string.
        :rtype: Color
        """
        return cls.colorize('bgred', string, auto=auto)

    @classmethod
    def green(cls, string, auto=False):
        """Color-code entire string.

        :param str string: String to colorize.
        :param bool auto: Enable auto-color (dark/light terminal).

        :return: Class instance for colorized string.
        :rtype: Color
        """
        return cls.colorize('green', string, auto=auto)

    @classmethod
    def bggreen(cls, string, auto=False):
        """Color-code entire string.

        :param str string: String to colorize.
        :param bool auto: Enable auto-color (dark/light terminal).

        :return: Class instance for colorized string.
        :rtype: Color
        """
        return cls.colorize('bggreen', string, auto=auto)

    @classmethod
    def yellow(cls, string, auto=False):
        """Color-code entire string.

        :param str string: String to colorize.
        :param bool auto: Enable auto-color (dark/light terminal).

        :return: Class instance for colorized string.
        :rtype: Color
        """
        return cls.colorize('yellow', string, auto=auto)

    @classmethod
    def bgyellow(cls, string, auto=False):
        """Color-code entire string.

        :param str string: String to colorize.
        :param bool auto: Enable auto-color (dark/light terminal).

        :return: Class instance for colorized string.
        :rtype: Color
        """
        return cls.colorize('bgyellow', string, auto=auto)

    @classmethod
    def blue(cls, string, auto=False):
        """Color-code entire string.

        :param str string: String to colorize.
        :param bool auto: Enable auto-color (dark/light terminal).

        :return: Class instance for colorized string.
        :rtype: Color
        """
        return cls.colorize('blue', string, auto=auto)

    @classmethod
    def bgblue(cls, string, auto=False):
        """Color-code entire string.

        :param str string: String to colorize.
        :param bool auto: Enable auto-color (dark/light terminal).

        :return: Class instance for colorized string.
        :rtype: Color
        """
        return cls.colorize('bgblue', string, auto=auto)

    @classmethod
    def magenta(cls, string, auto=False):
        """Color-code entire string.

        :param str string: String to colorize.
        :param bool auto: Enable auto-color (dark/light terminal).

        :return: Class instance for colorized string.
        :rtype: Color
        """
        return cls.colorize('magenta', string, auto=auto)

    @classmethod
    def bgmagenta(cls, string, auto=False):
        """Color-code entire string.

        :param str string: String to colorize.
        :param bool auto: Enable auto-color (dark/light terminal).

        :return: Class instance for colorized string.
        :rtype: Color
        """
        return cls.colorize('bgmagenta', string, auto=auto)

    @classmethod
    def cyan(cls, string, auto=False):
        """Color-code entire string.

        :param str string: String to colorize.
        :param bool auto: Enable auto-color (dark/light terminal).

        :return: Class instance for colorized string.
        :rtype: Color
        """
        return cls.colorize('cyan', string, auto=auto)

    @classmethod
    def bgcyan(cls, string, auto=False):
        """Color-code entire string.

        :param str string: String to colorize.
        :param bool auto: Enable auto-color (dark/light terminal).

        :return: Class instance for colorized string.
        :rtype: Color
        """
        return cls.colorize('bgcyan', string, auto=auto)

    @classmethod
    def white(cls, string, auto=False):
        """Color-code entire string.

        :param str string: String to colorize.
        :param bool auto: Enable auto-color (dark/light terminal).

        :return: Class instance for colorized string.
        :rtype: Color
        """
        return cls.colorize('white', string, auto=auto)

    @classmethod
    def bgwhite(cls, string, auto=False):
        """Color-code entire string.

        :param str string: String to colorize.
        :param bool auto: Enable auto-color (dark/light terminal).

        :return: Class instance for colorized string.
        :rtype: Color
        """
        return cls.colorize('bgwhite', string, auto=auto)
