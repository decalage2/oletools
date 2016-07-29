"""Determine color of characters that may or may not be adjacent to ANSI escape sequences."""

from colorclass.parse import RE_SPLIT


def build_color_index(ansi_string):
    """Build an index between visible characters and a string with invisible color codes.

    :param str ansi_string: String with color codes (ANSI escape sequences).

    :return: Position of visible characters in color string (indexes match non-color string).
    :rtype: tuple
    """
    mapping = list()
    color_offset = 0
    for item in (i for i in RE_SPLIT.split(ansi_string) if i):
        if RE_SPLIT.match(item):
            color_offset += len(item)
        else:
            for _ in range(len(item)):
                mapping.append(color_offset)
                color_offset += 1
    return tuple(mapping)


def find_char_color(ansi_string, pos):
    """Determine what color a character is in the string.

    :param str ansi_string: String with color codes (ANSI escape sequences).
    :param int pos: Position of the character in the ansi_string.

    :return: Character along with all surrounding color codes.
    :rtype: str
    """
    result = list()
    position = 0  # Set to None when character is found.
    for item in (i for i in RE_SPLIT.split(ansi_string) if i):
        if RE_SPLIT.match(item):
            result.append(item)
            if position is not None:
                position += len(item)
        elif position is not None:
            for char in item:
                if position == pos:
                    result.append(char)
                    position = None
                    break
                position += 1
    return ''.join(result)
