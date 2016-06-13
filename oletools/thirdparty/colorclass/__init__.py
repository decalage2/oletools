"""Colorful worry-free console applications for Linux, Mac OS X, and Windows.

Supported natively on Linux and Mac OSX (Just Works), and on Windows it works the same if Windows.enable() is called.

Gives you expected and sane results from methods like len() and .capitalize().

https://github.com/Robpol86/colorclass
https://pypi.python.org/pypi/colorclass
"""

from .codes import list_tags  # noqa
from .color import Color  # noqa
from .toggles import disable_all_colors  # noqa
from .toggles import disable_if_no_tty  # noqa
from .toggles import enable_all_colors  # noqa
from .toggles import is_enabled  # noqa
from .toggles import is_light  # noqa
from .toggles import set_dark_background  # noqa
from .toggles import set_light_background  # noqa
from .windows import Windows  # noqa


__all__ = (
    'Color',
    'disable_all_colors',
    'enable_all_colors',
    'is_enabled',
    'is_light',
    'list_tags',
    'set_dark_background',
    'set_light_background',
    'Windows',
)


__author__ = '@Robpol86'
__license__ = 'MIT'
__version__ = '2.2.0'
