"""Handles mapping between color names and ANSI codes and determining auto color codes."""

import sys
from collections import Mapping

BASE_CODES = {
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


class ANSICodeMapping(Mapping):
    """Read-only dictionary, resolves closing tags and automatic colors. Iterates only used color tags.

    :cvar bool DISABLE_COLORS: Disable colors (strip color codes).
    :cvar bool LIGHT_BACKGROUND: Use low intensity color codes.
    """

    DISABLE_COLORS = False
    LIGHT_BACKGROUND = False

    def __init__(self, value_markup):
        """Constructor.

        :param str value_markup: String with {color} tags.
        """
        self.whitelist = [k for k in BASE_CODES if '{' + k + '}' in value_markup]

    def __getitem__(self, item):
        """Return value for key or None if colors are disabled.

        :param str item: Key.

        :return: Color code integer.
        :rtype: int
        """
        if item not in self.whitelist:
            raise KeyError(item)
        if self.DISABLE_COLORS:
            return None
        return getattr(self, item, BASE_CODES[item])

    def __iter__(self):
        """Iterate dictionary."""
        return iter(self.whitelist)

    def __len__(self):
        """Dictionary length."""
        return len(self.whitelist)

    @classmethod
    def disable_all_colors(cls):
        """Disable all colors. Strips any color tags or codes."""
        cls.DISABLE_COLORS = True

    @classmethod
    def enable_all_colors(cls):
        """Enable all colors. Strips any color tags or codes."""
        cls.DISABLE_COLORS = False

    @classmethod
    def disable_if_no_tty(cls):
        """Disable all colors only if there is no TTY available.

        :return: True if colors are disabled, False if stderr or stdout is a TTY.
        :rtype: bool
        """
        if sys.stdout.isatty() or sys.stderr.isatty():
            return False
        cls.disable_all_colors()
        return True

    @classmethod
    def set_dark_background(cls):
        """Choose dark colors for all 'auto'-prefixed codes for readability on light backgrounds."""
        cls.LIGHT_BACKGROUND = False

    @classmethod
    def set_light_background(cls):
        """Choose dark colors for all 'auto'-prefixed codes for readability on light backgrounds."""
        cls.LIGHT_BACKGROUND = True

    @property
    def autoblack(self):
        """Return automatic black foreground color depending on background color."""
        return BASE_CODES['black' if ANSICodeMapping.LIGHT_BACKGROUND else 'hiblack']

    @property
    def autored(self):
        """Return automatic red foreground color depending on background color."""
        return BASE_CODES['red' if ANSICodeMapping.LIGHT_BACKGROUND else 'hired']

    @property
    def autogreen(self):
        """Return automatic green foreground color depending on background color."""
        return BASE_CODES['green' if ANSICodeMapping.LIGHT_BACKGROUND else 'higreen']

    @property
    def autoyellow(self):
        """Return automatic yellow foreground color depending on background color."""
        return BASE_CODES['yellow' if ANSICodeMapping.LIGHT_BACKGROUND else 'hiyellow']

    @property
    def autoblue(self):
        """Return automatic blue foreground color depending on background color."""
        return BASE_CODES['blue' if ANSICodeMapping.LIGHT_BACKGROUND else 'hiblue']

    @property
    def automagenta(self):
        """Return automatic magenta foreground color depending on background color."""
        return BASE_CODES['magenta' if ANSICodeMapping.LIGHT_BACKGROUND else 'himagenta']

    @property
    def autocyan(self):
        """Return automatic cyan foreground color depending on background color."""
        return BASE_CODES['cyan' if ANSICodeMapping.LIGHT_BACKGROUND else 'hicyan']

    @property
    def autowhite(self):
        """Return automatic white foreground color depending on background color."""
        return BASE_CODES['white' if ANSICodeMapping.LIGHT_BACKGROUND else 'hiwhite']

    @property
    def autobgblack(self):
        """Return automatic black background color depending on background color."""
        return BASE_CODES['bgblack' if ANSICodeMapping.LIGHT_BACKGROUND else 'hibgblack']

    @property
    def autobgred(self):
        """Return automatic red background color depending on background color."""
        return BASE_CODES['bgred' if ANSICodeMapping.LIGHT_BACKGROUND else 'hibgred']

    @property
    def autobggreen(self):
        """Return automatic green background color depending on background color."""
        return BASE_CODES['bggreen' if ANSICodeMapping.LIGHT_BACKGROUND else 'hibggreen']

    @property
    def autobgyellow(self):
        """Return automatic yellow background color depending on background color."""
        return BASE_CODES['bgyellow' if ANSICodeMapping.LIGHT_BACKGROUND else 'hibgyellow']

    @property
    def autobgblue(self):
        """Return automatic blue background color depending on background color."""
        return BASE_CODES['bgblue' if ANSICodeMapping.LIGHT_BACKGROUND else 'hibgblue']

    @property
    def autobgmagenta(self):
        """Return automatic magenta background color depending on background color."""
        return BASE_CODES['bgmagenta' if ANSICodeMapping.LIGHT_BACKGROUND else 'hibgmagenta']

    @property
    def autobgcyan(self):
        """Return automatic cyan background color depending on background color."""
        return BASE_CODES['bgcyan' if ANSICodeMapping.LIGHT_BACKGROUND else 'hibgcyan']

    @property
    def autobgwhite(self):
        """Return automatic white background color depending on background color."""
        return BASE_CODES['bgwhite' if ANSICodeMapping.LIGHT_BACKGROUND else 'hibgwhite']


def list_tags():
    """List the available tags.

    :return: List of 4-item tuples: opening tag, closing tag, main ansi value, closing ansi value.
    :rtype: list
    """
    # Build reverse dictionary. Keys are closing tags, values are [closing ansi, opening tag, opening ansi].
    reverse_dict = dict()
    for tag, ansi in sorted(BASE_CODES.items()):
        if tag.startswith('/'):
            reverse_dict[tag] = [ansi, None, None]
        else:
            reverse_dict['/' + tag][1:] = [tag, ansi]

    # Collapse
    four_item_tuples = [(v[1], k, v[2], v[0]) for k, v in reverse_dict.items()]

    # Sort.
    def sorter(four_item):
        """Sort /all /fg /bg first, then b i u flash, then auto colors, then dark colors, finally light colors.

        :param iter four_item: [opening tag, closing tag, main ansi value, closing ansi value]

        :return Sorting weight.
        :rtype: int
        """
        if not four_item[2]:  # /all /fg /bg
            return four_item[3] - 200
        if four_item[2] < 10 or four_item[0].startswith('auto'):  # b f i u or auto colors
            return four_item[2] - 100
        return four_item[2]
    four_item_tuples.sort(key=sorter)

    return four_item_tuples
