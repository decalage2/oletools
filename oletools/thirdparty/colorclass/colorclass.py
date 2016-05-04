"""Colorful worry-free console applications for Linux, Mac OS X, and Windows.

Supported natively on Linux and Mac OSX (Just Works), and on Windows it works the same if Windows.enable() is called.

Gives you expected and sane results from methods like len() and .capitalize().

https://github.com/Robpol86/colorclass
https://pypi.python.org/pypi/colorclass
"""

import atexit
from collections import Mapping
import ctypes
import os
import re
import sys

if os.name == 'nt':
    import ctypes.wintypes

__author__ = '@Robpol86'
__license__ = 'MIT'
__version__ = '1.2.0'
_BASE_CODES = {
    '/all': 0, 'b': 1, 'f': 2, 'i': 3, 'u': 4, 'flash': 5, 'outline': 6, 'negative': 7, 'invis': 8, 'strike': 9,
    '/b': 22, '/f': 22, '/i': 23, '/u': 24, '/flash': 25, '/outline': 26, '/negative': 27, '/invis': 28,
    '/strike': 29, '/fg': 39, '/bg': 49,

    'black': 30, 'red': 31, 'green': 32, 'yellow': 33, 'blue': 34, 'magenta': 35, 'cyan': 36, 'white': 37,

    'bgblack': 40, 'bgred': 41, 'bggreen': 42, 'bgyellow': 43, 'bgblue': 44, 'bgmagenta': 45, 'bgcyan': 46,
    'bgwhite': 47,

    'hiblack': 90, 'hired': 91, 'higreen': 92, 'hiyellow': 93, 'hiblue': 94, 'himagenta': 95, 'hicyan': 96,
    'hiwhite': 97,

    'hibgblack': 100, 'hibgred': 101, 'hibggreen': 102, 'hibgyellow': 103, 'hibgblue': 104, 'hibgmagenta': 105,
    'hibgcyan': 106, 'hibgwhite': 107,

    'autored': None, 'autoblack': None, 'automagenta': None, 'autowhite': None, 'autoblue': None, 'autoyellow': None,
    'autogreen': None, 'autocyan': None,

    'autobgred': None, 'autobgblack': None, 'autobgmagenta': None, 'autobgwhite': None, 'autobgblue': None,
    'autobgyellow': None, 'autobggreen': None, 'autobgcyan': None,

    '/black': 39, '/red': 39, '/green': 39, '/yellow': 39, '/blue': 39, '/magenta': 39, '/cyan': 39, '/white': 39,
    '/hiblack': 39, '/hired': 39, '/higreen': 39, '/hiyellow': 39, '/hiblue': 39, '/himagenta': 39, '/hicyan': 39,
    '/hiwhite': 39,

    '/bgblack': 49, '/bgred': 49, '/bggreen': 49, '/bgyellow': 49, '/bgblue': 49, '/bgmagenta': 49, '/bgcyan': 49,
    '/bgwhite': 49, '/hibgblack': 49, '/hibgred': 49, '/hibggreen': 49, '/hibgyellow': 49, '/hibgblue': 49,
    '/hibgmagenta': 49, '/hibgcyan': 49, '/hibgwhite': 49,

    '/autored': 39, '/autoblack': 39, '/automagenta': 39, '/autowhite': 39, '/autoblue': 39, '/autoyellow': 39,
    '/autogreen': 39, '/autocyan': 39,

    '/autobgred': 49, '/autobgblack': 49, '/autobgmagenta': 49, '/autobgwhite': 49, '/autobgblue': 49,
    '/autobgyellow': 49, '/autobggreen': 49, '/autobgcyan': 49,
}
_WINDOWS_CODES = {
    '/all': -33, '/fg': -39, '/bg': -49,

    'black': 0, 'red': 4, 'green': 2, 'yellow': 6, 'blue': 1, 'magenta': 5, 'cyan': 3, 'white': 7,

    'bgblack': -8, 'bgred': 64, 'bggreen': 32, 'bgyellow': 96, 'bgblue': 16, 'bgmagenta': 80, 'bgcyan': 48,
    'bgwhite': 112,

    'hiblack': 8, 'hired': 12, 'higreen': 10, 'hiyellow': 14, 'hiblue': 9, 'himagenta': 13, 'hicyan': 11, 'hiwhite': 15,

    'hibgblack': 128, 'hibgred': 192, 'hibggreen': 160, 'hibgyellow': 224, 'hibgblue': 144, 'hibgmagenta': 208,
    'hibgcyan': 176, 'hibgwhite': 240,

    '/black': -39, '/red': -39, '/green': -39, '/yellow': -39, '/blue': -39, '/magenta': -39, '/cyan': -39,
    '/white': -39, '/hiblack': -39, '/hired': -39, '/higreen': -39, '/hiyellow': -39, '/hiblue': -39, '/himagenta': -39,
    '/hicyan': -39, '/hiwhite': -39,

    '/bgblack': -49, '/bgred': -49, '/bggreen': -49, '/bgyellow': -49, '/bgblue': -49, '/bgmagenta': -49,
    '/bgcyan': -49, '/bgwhite': -49, '/hibgblack': -49, '/hibgred': -49, '/hibggreen': -49, '/hibgyellow': -49,
    '/hibgblue': -49, '/hibgmagenta': -49, '/hibgcyan': -49, '/hibgwhite': -49,
}
_RE_GROUP_SEARCH = re.compile(r'(?:\033\[[\d;]+m)+')
_RE_NUMBER_SEARCH = re.compile(r'\033\[([\d;]+)m')
_RE_SPLIT = re.compile(r'(\033\[[\d;]+m)')
PARENT_CLASS = type(u'')


class _AutoCodes(Mapping):
    """Read-only subclass of dict, resolves closing tags (based on colorclass.CODES) and automatic colors."""
    DISABLE_COLORS = False
    LIGHT_BACKGROUND = False

    def __init__(self):
        self.__dict = _BASE_CODES.copy()

    def __getitem__(self, item):
        if item == 'autoblack':
            answer = self.autoblack
        elif item == 'autored':
            answer = self.autored
        elif item == 'autogreen':
            answer = self.autogreen
        elif item == 'autoyellow':
            answer = self.autoyellow
        elif item == 'autoblue':
            answer = self.autoblue
        elif item == 'automagenta':
            answer = self.automagenta
        elif item == 'autocyan':
            answer = self.autocyan
        elif item == 'autowhite':
            answer = self.autowhite
        elif item == 'autobgblack':
            answer = self.autobgblack
        elif item == 'autobgred':
            answer = self.autobgred
        elif item == 'autobggreen':
            answer = self.autobggreen
        elif item == 'autobgyellow':
            answer = self.autobgyellow
        elif item == 'autobgblue':
            answer = self.autobgblue
        elif item == 'autobgmagenta':
            answer = self.autobgmagenta
        elif item == 'autobgcyan':
            answer = self.autobgcyan
        elif item == 'autobgwhite':
            answer = self.autobgwhite
        else:
            answer = self.__dict[item]
        return answer

    def __iter__(self):
        return iter(self.__dict)

    def __len__(self):
        return len(self.__dict)

    @property
    def autoblack(self):
        """Returns automatic black foreground color depending on background color."""
        return self.__dict['black' if _AutoCodes.LIGHT_BACKGROUND else 'hiblack']

    @property
    def autored(self):
        """Returns automatic red foreground color depending on background color."""
        return self.__dict['red' if _AutoCodes.LIGHT_BACKGROUND else 'hired']

    @property
    def autogreen(self):
        """Returns automatic green foreground color depending on background color."""
        return self.__dict['green' if _AutoCodes.LIGHT_BACKGROUND else 'higreen']

    @property
    def autoyellow(self):
        """Returns automatic yellow foreground color depending on background color."""
        return self.__dict['yellow' if _AutoCodes.LIGHT_BACKGROUND else 'hiyellow']

    @property
    def autoblue(self):
        """Returns automatic blue foreground color depending on background color."""
        return self.__dict['blue' if _AutoCodes.LIGHT_BACKGROUND else 'hiblue']

    @property
    def automagenta(self):
        """Returns automatic magenta foreground color depending on background color."""
        return self.__dict['magenta' if _AutoCodes.LIGHT_BACKGROUND else 'himagenta']

    @property
    def autocyan(self):
        """Returns automatic cyan foreground color depending on background color."""
        return self.__dict['cyan' if _AutoCodes.LIGHT_BACKGROUND else 'hicyan']

    @property
    def autowhite(self):
        """Returns automatic white foreground color depending on background color."""
        return self.__dict['white' if _AutoCodes.LIGHT_BACKGROUND else 'hiwhite']

    @property
    def autobgblack(self):
        """Returns automatic black background color depending on background color."""
        return self.__dict['bgblack' if _AutoCodes.LIGHT_BACKGROUND else 'hibgblack']

    @property
    def autobgred(self):
        """Returns automatic red background color depending on background color."""
        return self.__dict['bgred' if _AutoCodes.LIGHT_BACKGROUND else 'hibgred']

    @property
    def autobggreen(self):
        """Returns automatic green background color depending on background color."""
        return self.__dict['bggreen' if _AutoCodes.LIGHT_BACKGROUND else 'hibggreen']

    @property
    def autobgyellow(self):
        """Returns automatic yellow background color depending on background color."""
        return self.__dict['bgyellow' if _AutoCodes.LIGHT_BACKGROUND else 'hibgyellow']

    @property
    def autobgblue(self):
        """Returns automatic blue background color depending on background color."""
        return self.__dict['bgblue' if _AutoCodes.LIGHT_BACKGROUND else 'hibgblue']

    @property
    def autobgmagenta(self):
        """Returns automatic magenta background color depending on background color."""
        return self.__dict['bgmagenta' if _AutoCodes.LIGHT_BACKGROUND else 'hibgmagenta']

    @property
    def autobgcyan(self):
        """Returns automatic cyan background color depending on background color."""
        return self.__dict['bgcyan' if _AutoCodes.LIGHT_BACKGROUND else 'hibgcyan']

    @property
    def autobgwhite(self):
        """Returns automatic white background color depending on background color."""
        return self.__dict['bgwhite' if _AutoCodes.LIGHT_BACKGROUND else 'hibgwhite']


def _pad_input(incoming):
    """Avoid IndexError and KeyError by ignoring un-related fields.

    Example: '{0}{autored}' becomes '{{0}}{autored}'.

    Positional arguments:
    incoming -- the input unicode value.

    Returns:
    Padded unicode value.
    """
    incoming_expanded = incoming.replace('{', '{{').replace('}', '}}')
    for key in _BASE_CODES:
        before, after = '{{%s}}' % key, '{%s}' % key
        if before in incoming_expanded:
            incoming_expanded = incoming_expanded.replace(before, after)
    return incoming_expanded


def _parse_input(incoming):
    """Performs the actual conversion of tags to ANSI escaped codes.

    Provides a version of the input without any colors for len() and other methods.

    Positional arguments:
    incoming -- the input unicode value.

    Returns:
    2-item tuple. First item is the parsed output. Second item is a version of the input without any colors.
    """
    codes = dict((k, v) for k, v in _AutoCodes().items() if '{%s}' % k in incoming)
    color_codes = dict((k, '' if _AutoCodes.DISABLE_COLORS else '\033[{0}m'.format(v)) for k, v in codes.items())
    incoming_padded = _pad_input(incoming)
    output_colors = incoming_padded.format(**color_codes)

    # Simplify: '{b}{red}' -> '\033[1m\033[31m' -> '\033[1;31m'
    groups = sorted(set(_RE_GROUP_SEARCH.findall(output_colors)), key=len, reverse=True)  # Get codes, grouped adjacent.
    groups_simplified = [[x for n in _RE_NUMBER_SEARCH.findall(i) for x in n.split(';')] for i in groups]
    groups_compiled = ['\033[{0}m'.format(';'.join(g)) for g in groups_simplified]  # Final codes.
    assert len(groups_compiled) == len(groups)  # For testing.
    output_colors_simplified = output_colors
    for i in range(len(groups)):
        output_colors_simplified = output_colors_simplified.replace(groups[i], groups_compiled[i])
    output_no_colors = _RE_SPLIT.sub('', output_colors_simplified)

    # Strip any remaining color codes.
    if _AutoCodes.DISABLE_COLORS:
        output_colors_simplified = _RE_NUMBER_SEARCH.sub('', output_colors_simplified)

    return output_colors_simplified, output_no_colors


def disable_all_colors():
    """Disable all colors. Strips any color tags or codes."""
    _AutoCodes.DISABLE_COLORS = True


def set_light_background():
    """Chooses dark colors for all 'auto'-prefixed codes for readability on light backgrounds."""
    _AutoCodes.DISABLE_COLORS = False
    _AutoCodes.LIGHT_BACKGROUND = True


def set_dark_background():
    """Chooses dark colors for all 'auto'-prefixed codes for readability on light backgrounds."""
    _AutoCodes.DISABLE_COLORS = False
    _AutoCodes.LIGHT_BACKGROUND = False


def list_tags():
    """Lists the available tags.

    Returns:
    Tuple of tuples. Child tuples are four items: ('opening tag', 'closing tag', main ansi value, closing ansi value).
    """
    codes = _AutoCodes()
    grouped = set([(k, '/{0}'.format(k), codes[k], codes['/{0}'.format(k)]) for k in codes if not k.startswith('/')])

    # Add half-tags like /all.
    found = [c for r in grouped for c in r[:2]]
    missing = set([('', r[0], None, r[1]) if r[0].startswith('/') else (r[0], '', r[1], None)
                   for r in _AutoCodes().items() if r[0] not in found])
    grouped |= missing

    # Sort.
    payload = sorted([i for i in grouped if i[2] is None], key=lambda x: x[3])  # /all /fg /bg
    grouped -= set(payload)
    payload.extend(sorted([i for i in grouped if i[2] < 10], key=lambda x: x[2]))  # b i u flash
    grouped -= set(payload)
    payload.extend(sorted([i for i in grouped if i[0].startswith('auto')], key=lambda x: x[2]))  # auto colors
    grouped -= set(payload)
    payload.extend(sorted([i for i in grouped if not i[0].startswith('hi')], key=lambda x: x[2]))  # dark colors
    grouped -= set(payload)
    payload.extend(sorted(grouped, key=lambda x: x[2]))  # light colors
    return tuple(payload)


class ColorBytes(bytes):
    """Str (bytes in Python3) subclass, .decode() overridden to return Color() instance."""

    def decode(*args, **kwargs):
        return Color(super(ColorBytes, args[0]).decode(*args[1:], **kwargs))


class Color(PARENT_CLASS):
    """Unicode (str in Python3) subclass with ANSI terminal text color support.

    Example syntax: Color('{red}Sample Text{/red}')

    For a list of codes, call: colorclass.list_tags()
    """

    @classmethod
    def red(cls, s, auto=False):
        return cls.colorize('red', s, auto=auto)

    @classmethod
    def bgred(cls, s, auto=False):
        return cls.colorize('bgred', s, auto=auto)

    @classmethod
    def green(cls, s, auto=False):
        return cls.colorize('green', s, auto=auto)

    @classmethod
    def bggreen(cls, s, auto=False):
        return cls.colorize('bggreen', s, auto=auto)

    @classmethod
    def blue(cls, s, auto=False):
        return cls.colorize('blue', s, auto=auto)

    @classmethod
    def bgblue(cls, s, auto=False):
        return cls.colorize('bgblue', s, auto=auto)

    @classmethod
    def yellow(cls, s, auto=False):
        return cls.colorize('yellow', s, auto=auto)

    @classmethod
    def bgyellow(cls, s, auto=False):
        return cls.colorize('bgyellow', s, auto=auto)

    @classmethod
    def cyan(cls, s, auto=False):
        return cls.colorize('cyan', s, auto=auto)

    @classmethod
    def bgcyan(cls, s, auto=False):
        return cls.colorize('bgcyan', s, auto=auto)

    @classmethod
    def magenta(cls, s, auto=False):
        return cls.colorize('magenta', s, auto=auto)

    @classmethod
    def bgmagenta(cls, s, auto=False):
        return cls.colorize('bgmagenta', s, auto=auto)

    @classmethod
    def colorize(cls, color, s, auto=False):
        tag = '{0}{1}'.format('auto' if auto else '', color)
        return cls('{%s}%s{/%s}' % (tag, s, tag))

    def __new__(cls, *args, **kwargs):
        parent_class = cls.__bases__[0]
        value_markup = args[0] if args else parent_class()
        value_colors, value_no_colors = _parse_input(value_markup)
        if args:
            args = [value_colors] + list(args[1:])

        obj = parent_class.__new__(cls, *args, **kwargs)
        obj.value_colors, obj.value_no_colors = value_colors, value_no_colors
        obj.has_colors = bool(_RE_NUMBER_SEARCH.match(value_colors))
        return obj

    def __len__(self):
        return self.value_no_colors.__len__()

    def capitalize(self):
        split = _RE_SPLIT.split(self.value_colors)
        for i in range(len(split)):
            if _RE_SPLIT.match(split[i]):
                continue
            split[i] = PARENT_CLASS(split[i]).capitalize()
        return Color().join(split)

    def center(self, width, fillchar=None):
        if fillchar is not None:
            result = PARENT_CLASS(self.value_no_colors).center(width, fillchar)
        else:
            result = PARENT_CLASS(self.value_no_colors).center(width)
        return result.replace(self.value_no_colors, self.value_colors)

    def count(self, *args, **kwargs):
        return PARENT_CLASS(self.value_no_colors).count(*args, **kwargs)

    def endswith(self, *args, **kwargs):
        return PARENT_CLASS(self.value_no_colors).endswith(*args, **kwargs)

    def encode(*args, **kwargs):
        return ColorBytes(super(Color, args[0]).encode(*args[1:], **kwargs))

    def decode(*args, **kwargs):
        return Color(super(Color, args[0]).decode(*args[1:], **kwargs))

    def find(self, *args, **kwargs):
        return PARENT_CLASS(self.value_no_colors).find(*args, **kwargs)

    def format(*args, **kwargs):
        return Color(super(Color, args[0]).format(*args[1:], **kwargs))

    def index(self, *args, **kwargs):
        return PARENT_CLASS(self.value_no_colors).index(*args, **kwargs)

    def isalnum(self):
        return PARENT_CLASS(self.value_no_colors).isalnum()

    def isalpha(self):
        return PARENT_CLASS(self.value_no_colors).isalpha()

    def isdecimal(self):
        return PARENT_CLASS(self.value_no_colors).isdecimal()

    def isdigit(self):
        return PARENT_CLASS(self.value_no_colors).isdigit()

    def isnumeric(self):
        return PARENT_CLASS(self.value_no_colors).isnumeric()

    def isspace(self):
        return PARENT_CLASS(self.value_no_colors).isspace()

    def istitle(self):
        return PARENT_CLASS(self.value_no_colors).istitle()

    def isupper(self):
        return PARENT_CLASS(self.value_no_colors).isupper()

    def ljust(self, width, fillchar=None):
        if fillchar is not None:
            result = PARENT_CLASS(self.value_no_colors).ljust(width, fillchar)
        else:
            result = PARENT_CLASS(self.value_no_colors).ljust(width)
        return result.replace(self.value_no_colors, self.value_colors)

    def rfind(self, *args, **kwargs):
        return PARENT_CLASS(self.value_no_colors).rfind(*args, **kwargs)

    def rindex(self, *args, **kwargs):
        return PARENT_CLASS(self.value_no_colors).rindex(*args, **kwargs)

    def rjust(self, width, fillchar=None):
        if fillchar is not None:
            result = PARENT_CLASS(self.value_no_colors).rjust(width, fillchar)
        else:
            result = PARENT_CLASS(self.value_no_colors).rjust(width)
        return result.replace(self.value_no_colors, self.value_colors)

    def splitlines(self):
        return [Color(l) for l in PARENT_CLASS(self.value_colors).splitlines()]

    def startswith(self, *args, **kwargs):
        return PARENT_CLASS(self.value_no_colors).startswith(*args, **kwargs)

    def swapcase(self):
        split = _RE_SPLIT.split(self.value_colors)
        for i in range(len(split)):
            if _RE_SPLIT.match(split[i]):
                continue
            split[i] = PARENT_CLASS(split[i]).swapcase()
        return Color().join(split)

    def title(self):
        split = _RE_SPLIT.split(self.value_colors)
        for i in range(len(split)):
            if _RE_SPLIT.match(split[i]):
                continue
            split[i] = PARENT_CLASS(split[i]).title()
        return Color().join(split)

    def translate(self, table):
        split = _RE_SPLIT.split(self.value_colors)
        for i in range(len(split)):
            if _RE_SPLIT.match(split[i]):
                continue
            split[i] = PARENT_CLASS(split[i]).translate(table)
        return Color().join(split)

    def upper(self):
        split = _RE_SPLIT.split(self.value_colors)
        for i in range(len(split)):
            if _RE_SPLIT.match(split[i]):
                continue
            split[i] = PARENT_CLASS(split[i]).upper()
        return Color().join(split)

    def zfill(self, width):
        if not self.value_no_colors:
            return PARENT_CLASS().zfill(width)

        split = _RE_SPLIT.split(self.value_colors)
        filled = PARENT_CLASS(self.value_no_colors).zfill(width)
        if len(split) == 1:
            return filled

        padding = filled.replace(self.value_no_colors, '')
        if not split[0]:
            split[2] = padding + split[2]
        else:
            split[0] = padding + split[0]

        return Color().join(split)


class Windows(object):
    """Enable and disable Windows support for ANSI color character codes.

    Call static method Windows.enable() to enable color support for the remainder of the process' lifetime.

    This class is also a context manager. You can do this:
    with Windows():
        print(Color('{autored}Test{/autored}'))

    Or this:
    with Windows(auto_colors=True):
        print(Color('{autored}Test{/autored}'))
    """

    @staticmethod
    def disable():
        """Restore sys.stderr and sys.stdout to their original objects. Resets colors to their original values."""
        if os.name != 'nt' or not Windows.is_enabled():
            return False

        getattr(sys.stderr, '_reset_colors', lambda: False)()
        getattr(sys.stdout, '_reset_colors', lambda: False)()

        if hasattr(sys.stderr, 'ORIGINAL_STREAM'):
            sys.stderr = getattr(sys.stderr, 'ORIGINAL_STREAM')
        if hasattr(sys.stdout, 'ORIGINAL_STREAM'):
            sys.stdout = getattr(sys.stdout, 'ORIGINAL_STREAM')

        return True

    @staticmethod
    def is_enabled():
        """Returns True if either stderr or stdout has colors enabled."""
        return hasattr(sys.stderr, 'ORIGINAL_STREAM') or hasattr(sys.stdout, 'ORIGINAL_STREAM')

    @staticmethod
    def enable(auto_colors=False, reset_atexit=False):
        """Enables color text with print() or sys.stdout.write() (stderr too).

        Keyword arguments:
        auto_colors -- automatically selects dark or light colors based on current terminal's background color. Only
            works with {autored} and related tags.
        reset_atexit -- resets original colors upon Python exit (in case you forget to reset it yourself with a closing
            tag).
        """
        if os.name != 'nt':
            return False

        # Overwrite stream references.
        if not hasattr(sys.stderr, 'ORIGINAL_STREAM'):
            sys.stderr.flush()
            sys.stderr = _WindowsStreamStdErr()
        if not hasattr(sys.stdout, 'ORIGINAL_STREAM'):
            sys.stdout.flush()
            sys.stdout = _WindowsStreamStdOut()
        if not hasattr(sys.stderr, 'ORIGINAL_STREAM') and not hasattr(sys.stdout, 'ORIGINAL_STREAM'):
            return False

        # Automatically select which colors to display.
        bg_color = getattr(sys.stdout, 'default_bg', getattr(sys.stderr, 'default_bg', None))
        if auto_colors and bg_color is not None:
            set_light_background() if bg_color in (112, 96, 240, 176, 224, 208, 160) else set_dark_background()

        # Reset on exit if requested.
        if reset_atexit:
            atexit.register(lambda: Windows.disable())

        return True

    def __init__(self, auto_colors=False):
        self.auto_colors = auto_colors

    def __enter__(self):
        Windows.enable(auto_colors=self.auto_colors)

    def __exit__(self, *_):
        Windows.disable()


class _WindowsCSBI(object):
    """Interfaces with Windows CONSOLE_SCREEN_BUFFER_INFO API/DLL calls. Gets info for stderr and stdout.

    References:
        http://msdn.microsoft.com/en-us/library/windows/desktop/ms683231
        https://code.google.com/p/colorama/issues/detail?id=47.
        pytest's py project: py/_io/terminalwriter.py.

    Class variables:
    CSBI -- ConsoleScreenBufferInfo class/struct (not instance, the class definition itself) defined in _define_csbi().
    HANDLE_STDERR -- GetStdHandle() return integer for stderr.
    HANDLE_STDOUT -- GetStdHandle() return integer for stdout.
    WINDLL -- my own loaded instance of ctypes.WinDLL.
    """

    CSBI = None
    HANDLE_STDERR = None
    HANDLE_STDOUT = None
    WINDLL = ctypes.LibraryLoader(getattr(ctypes, 'WinDLL', None))

    @staticmethod
    def _define_csbi():
        """Defines structs and populates _WindowsCSBI.CSBI."""
        if _WindowsCSBI.CSBI is not None:
            return

        class COORD(ctypes.Structure):
            """Windows COORD structure. http://msdn.microsoft.com/en-us/library/windows/desktop/ms682119"""
            _fields_ = [('X', ctypes.c_short), ('Y', ctypes.c_short)]

        class SmallRECT(ctypes.Structure):
            """Windows SMALL_RECT structure. http://msdn.microsoft.com/en-us/library/windows/desktop/ms686311"""
            _fields_ = [('Left', ctypes.c_short), ('Top', ctypes.c_short), ('Right', ctypes.c_short),
                        ('Bottom', ctypes.c_short)]

        class ConsoleScreenBufferInfo(ctypes.Structure):
            """Windows CONSOLE_SCREEN_BUFFER_INFO structure.
            http://msdn.microsoft.com/en-us/library/windows/desktop/ms682093
            """
            _fields_ = [
                ('dwSize', COORD),
                ('dwCursorPosition', COORD),
                ('wAttributes', ctypes.wintypes.WORD),
                ('srWindow', SmallRECT),
                ('dwMaximumWindowSize', COORD)
            ]

        _WindowsCSBI.CSBI = ConsoleScreenBufferInfo

    @staticmethod
    def initialize():
        """Initializes the WINDLL resource and populated the CSBI class variable."""
        _WindowsCSBI._define_csbi()
        _WindowsCSBI.HANDLE_STDERR = _WindowsCSBI.HANDLE_STDERR or _WindowsCSBI.WINDLL.kernel32.GetStdHandle(-12)
        _WindowsCSBI.HANDLE_STDOUT = _WindowsCSBI.HANDLE_STDOUT or _WindowsCSBI.WINDLL.kernel32.GetStdHandle(-11)
        if _WindowsCSBI.WINDLL.kernel32.GetConsoleScreenBufferInfo.argtypes:
            return

        _WindowsCSBI.WINDLL.kernel32.GetStdHandle.argtypes = [ctypes.wintypes.DWORD]
        _WindowsCSBI.WINDLL.kernel32.GetStdHandle.restype = ctypes.wintypes.HANDLE
        _WindowsCSBI.WINDLL.kernel32.GetConsoleScreenBufferInfo.restype = ctypes.wintypes.BOOL
        _WindowsCSBI.WINDLL.kernel32.GetConsoleScreenBufferInfo.argtypes = [
            ctypes.wintypes.HANDLE, ctypes.POINTER(_WindowsCSBI.CSBI)
        ]

    @staticmethod
    def get_info(handle):
        """Get information about this current console window (for Microsoft Windows only).

        Raises IOError if attempt to get information fails (if there is no console window).

        Don't forget to call _WindowsCSBI.initialize() once in your application before calling this method.

        Positional arguments:
        handle -- either _WindowsCSBI.HANDLE_STDERR or _WindowsCSBI.HANDLE_STDOUT.

        Returns:
        Dictionary with different integer values. Keys are:
            buffer_width -- width of the buffer (Screen Buffer Size in cmd.exe layout tab).
            buffer_height -- height of the buffer (Screen Buffer Size in cmd.exe layout tab).
            terminal_width -- width of the terminal window.
            terminal_height -- height of the terminal window.
            bg_color -- current background color (http://msdn.microsoft.com/en-us/library/windows/desktop/ms682088).
            fg_color -- current text color code.
        """
        # Query Win32 API.
        csbi = _WindowsCSBI.CSBI()
        try:
            if not _WindowsCSBI.WINDLL.kernel32.GetConsoleScreenBufferInfo(handle, ctypes.byref(csbi)):
                raise IOError('Unable to get console screen buffer info from win32 API.')
        except ctypes.ArgumentError:
            raise IOError('Unable to get console screen buffer info from win32 API.')

        # Parse data.
        result = dict(
            buffer_width=int(csbi.dwSize.X - 1),
            buffer_height=int(csbi.dwSize.Y),
            terminal_width=int(csbi.srWindow.Right - csbi.srWindow.Left),
            terminal_height=int(csbi.srWindow.Bottom - csbi.srWindow.Top),
            bg_color=int(csbi.wAttributes & 240),
            fg_color=int(csbi.wAttributes % 16),
        )
        return result


class _WindowsStreamStdOut(object):
    """Replacement stream which overrides sys.stdout. When writing or printing, ANSI codes are converted.

    ANSI (Linux/Unix) color codes are converted into win32 system calls, changing the next character's color before
    printing it. Resources referenced:
        https://github.com/tartley/colorama
        http://www.cplusplus.com/articles/2ywTURfi/
        http://thomasfischer.biz/python-and-windows-terminal-colors/
        http://stackoverflow.com/questions/17125440/c-win32-console-color
        http://www.tysos.org/svn/trunk/mono/corlib/System/WindowsConsoleDriver.cs
        http://stackoverflow.com/questions/287871/print-in-terminal-with-colors-using-python
        http://msdn.microsoft.com/en-us/library/windows/desktop/ms682088#_win32_character_attributes

    Class variables:
    ALL_BG_CODES -- list of background Windows codes. Used to determine if requested color is foreground or background.
    COMPILED_CODES -- 'translation' dictionary. Keys are ANSI codes (values of _BASE_CODES), values are Windows codes.
    ORIGINAL_STREAM -- the original stream to write non-code text to.
    WIN32_STREAM_HANDLE -- handle to the Windows stdout device. Used by other Windows functions.

    Instance variables:
    default_fg -- the foreground Windows color code at the time of instantiation.
    default_bg -- the background Windows color code at the time of instantiation.
    """

    ALL_BG_CODES = [v for k, v in _WINDOWS_CODES.items() if k.startswith('bg') or k.startswith('hibg')]
    COMPILED_CODES = dict((v, _WINDOWS_CODES[k]) for k, v in _BASE_CODES.items() if k in _WINDOWS_CODES)
    ORIGINAL_STREAM = sys.stdout
    WIN32_STREAM_HANDLE = _WindowsCSBI.HANDLE_STDOUT

    def __init__(self):
        _WindowsCSBI.initialize()
        self.default_fg, self.default_bg = self._get_colors()
        for attr in dir(self.ORIGINAL_STREAM):
            if hasattr(self, attr):
                continue
            setattr(self, attr, getattr(self.ORIGINAL_STREAM, attr))

    def __getattr__(self, item):
        """If an attribute/function/etc is not defined in this function, retrieve the one from the original stream.

        Fixes ipython arrow key presses.
        """
        return getattr(self.ORIGINAL_STREAM, item)

    def _get_colors(self):
        """Returns a tuple of two integers representing current colors: (foreground, background)."""
        try:
            csbi = _WindowsCSBI.get_info(self.WIN32_STREAM_HANDLE)
            return csbi['fg_color'], csbi['bg_color']
        except IOError:
            return 7, 0

    def _reset_colors(self):
        """Sets the foreground and background colors to their original values (when class was instantiated)."""
        self._set_color(-33)

    def _set_color(self, color_code):
        """Changes the foreground and background colors for subsequently printed characters.

        Since setting a color requires including both foreground and background codes (merged), setting just the
        foreground color resets the background color to black, and vice versa.

        This function first gets the current background and foreground colors, merges in the requested color code, and
        sets the result.

        However if we need to remove just the foreground color but leave the background color the same (or vice versa)
        such as when {/red} is used, we must merge the default foreground color with the current background color. This
        is the reason for those negative values.

        Positional arguments:
        color_code -- integer color code from _WINDOWS_CODES.
        """
        # Get current color code.
        current_fg, current_bg = self._get_colors()

        # Handle special negative codes. Also determine the final color code.
        if color_code == -39:
            final_color_code = self.default_fg | current_bg  # Reset the foreground only.
        elif color_code == -49:
            final_color_code = current_fg | self.default_bg  # Reset the background only.
        elif color_code == -33:
            final_color_code = self.default_fg | self.default_bg  # Reset both.
        elif color_code == -8:
            final_color_code = current_fg  # Black background.
        else:
            new_is_bg = color_code in self.ALL_BG_CODES
            final_color_code = color_code | (current_fg if new_is_bg else current_bg)

        # Set new code.
        _WindowsCSBI.WINDLL.kernel32.SetConsoleTextAttribute(self.WIN32_STREAM_HANDLE, final_color_code)

    def write(self, p_str):
        for segment in _RE_SPLIT.split(p_str):
            if not segment:
                # Empty string. p_str probably starts with colors so the first item is always ''.
                continue
            if not _RE_SPLIT.match(segment):
                # No color codes, print regular text.
                self.ORIGINAL_STREAM.write(segment)
                self.ORIGINAL_STREAM.flush()
                continue
            for color_code in (int(c) for c in _RE_NUMBER_SEARCH.findall(segment)[0].split(';')):
                if color_code in self.COMPILED_CODES:
                    self._set_color(self.COMPILED_CODES[color_code])


class _WindowsStreamStdErr(_WindowsStreamStdOut):
    """Replacement stream which overrides sys.stderr. Subclasses _WindowsStreamStdOut."""

    ORIGINAL_STREAM = sys.stderr
    WIN32_STREAM_HANDLE = _WindowsCSBI.HANDLE_STDERR
