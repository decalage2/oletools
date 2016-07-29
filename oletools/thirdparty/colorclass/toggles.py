"""Convenience functions to enable/disable features."""

from colorclass.codes import ANSICodeMapping


def disable_all_colors():
    """Disable all colors. Strip any color tags or codes."""
    ANSICodeMapping.disable_all_colors()


def enable_all_colors():
    """Enable colors."""
    ANSICodeMapping.enable_all_colors()


def disable_if_no_tty():
    """Disable all colors if there is no TTY available.

    :return: True if colors are disabled, False if stderr or stdout is a TTY.
    :rtype: bool
    """
    return ANSICodeMapping.disable_if_no_tty()


def is_enabled():
    """Are colors enabled."""
    return not ANSICodeMapping.DISABLE_COLORS


def set_light_background():
    """Choose dark colors for all 'auto'-prefixed codes for readability on light backgrounds."""
    ANSICodeMapping.set_light_background()


def set_dark_background():
    """Choose dark colors for all 'auto'-prefixed codes for readability on light backgrounds."""
    ANSICodeMapping.set_dark_background()


def is_light():
    """Are background colors for light backgrounds."""
    return ANSICodeMapping.LIGHT_BACKGROUND
