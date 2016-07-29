"""Parse color markup tags into ANSI escape sequences."""

import re

from colorclass.codes import ANSICodeMapping, BASE_CODES

CODE_GROUPS = (
    tuple(set(str(i) for i in BASE_CODES.values() if i and (40 <= i <= 49 or 100 <= i <= 109))),  # bg colors
    tuple(set(str(i) for i in BASE_CODES.values() if i and (30 <= i <= 39 or 90 <= i <= 99))),  # fg colors
    ('1', '22'), ('2', '22'), ('3', '23'), ('4', '24'), ('5', '25'), ('6', '26'), ('7', '27'), ('8', '28'), ('9', '29'),
)
RE_ANSI = re.compile(r'(\033\[([\d;]+)m)')
RE_COMBINE = re.compile(r'\033\[([\d;]+)m\033\[([\d;]+)m')
RE_SPLIT = re.compile(r'(\033\[[\d;]+m)')


def prune_overridden(ansi_string):
    """Remove color codes that are rendered ineffective by subsequent codes in one escape sequence then sort codes.

    :param str ansi_string: Incoming ansi_string with ANSI color codes.

    :return: Color string with pruned color sequences.
    :rtype: str
    """
    multi_seqs = set(p for p in RE_ANSI.findall(ansi_string) if ';' in p[1])  # Sequences with multiple color codes.

    for escape, codes in multi_seqs:
        r_codes = list(reversed(codes.split(';')))

        # Nuke everything before {/all}.
        try:
            r_codes = r_codes[:r_codes.index('0') + 1]
        except ValueError:
            pass

        # Thin out groups.
        for group in CODE_GROUPS:
            for pos in reversed([i for i, n in enumerate(r_codes) if n in group][1:]):
                r_codes.pop(pos)

        # Done.
        reduced_codes = ';'.join(sorted(r_codes, key=int))
        if codes != reduced_codes:
            ansi_string = ansi_string.replace(escape, '\033[' + reduced_codes + 'm')

    return ansi_string


def parse_input(tagged_string, disable_colors, keep_tags):
    """Perform the actual conversion of tags to ANSI escaped codes.

    Provides a version of the input without any colors for len() and other methods.

    :param str tagged_string: The input unicode value.
    :param bool disable_colors: Strip all colors in both outputs.
    :param bool keep_tags: Skip parsing curly bracket tags into ANSI escape sequences.

    :return: 2-item tuple. First item is the parsed output. Second item is a version of the input without any colors.
    :rtype: tuple
    """
    codes = ANSICodeMapping(tagged_string)
    output_colors = getattr(tagged_string, 'value_colors', tagged_string)

    # Convert: '{b}{red}' -> '\033[1m\033[31m'
    if not keep_tags:
        for tag, replacement in (('{' + k + '}', '' if v is None else '\033[%dm' % v) for k, v in codes.items()):
            output_colors = output_colors.replace(tag, replacement)

    # Strip colors.
    output_no_colors = RE_ANSI.sub('', output_colors)
    if disable_colors:
        return output_no_colors, output_no_colors

    # Combine: '\033[1m\033[31m' -> '\033[1;31m'
    while True:
        simplified = RE_COMBINE.sub(r'\033[\1;\2m', output_colors)
        if simplified == output_colors:
            break
        output_colors = simplified

    # Prune: '\033[31;32;33;34;35m' -> '\033[35m'
    output_colors = prune_overridden(output_colors)

    # Deduplicate: '\033[1;mT\033[1;mE\033[1;mS\033[1;mT' -> '\033[1;mTEST'
    previous_escape = None
    segments = list()
    for item in (i for i in RE_SPLIT.split(output_colors) if i):
        if RE_SPLIT.match(item):
            if item != previous_escape:
                segments.append(item)
                previous_escape = item
        else:
            segments.append(item)
    output_colors = ''.join(segments)

    return output_colors, output_no_colors
