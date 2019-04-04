"""
codepages.py

codepages is a python module to map code pages (numbers) to Python codecs,
in order to decode bytes to unicode.
It also provides the name/description of code pages.

Author: Philippe Lagadec - http://www.decalage.info
License: BSD, see source code or documentation

codepages is part of the python-oletools package:
http://www.decalage.info/python/oletools
"""

# === LICENSE ==================================================================

# codepages is copyright (c) 2018-2019 Philippe Lagadec (http://www.decalage.info)
# All rights reserved.
#
# Redistribution and use in source and binary forms, with or without modification,
# are permitted provided that the following conditions are met:
#
#  * Redistributions of source code must retain the above copyright notice, this
#    list of conditions and the following disclaimer.
#  * Redistributions in binary form must reproduce the above copyright notice,
#    this list of conditions and the following disclaimer in the documentation
#    and/or other materials provided with the distribution.
#
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
# ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
# WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
# FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
# DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
# SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
# CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
# OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
# OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.


# -----------------------------------------------------------------------------
# CHANGELOG:
# 2018-12-13 v0.54 PL: - first version
# 2019-01-30       PL: - added a few code pages from xlrd

__version__ = '0.54'

# -----------------------------------------------------------------------------
# TODO:
# TODO: check also http://www.aivosto.com/articles/charsets-codepages.html
# TODO: https://en.wikipedia.org/wiki/Code_page

# -----------------------------------------------------------------------------
# REFERENCES:
# - https://docs.microsoft.com/en-gb/windows/desktop/Intl/code-page-identifiers


# --- IMPORTS -----------------------------------------------------------------

import codecs

# === CONSTANTS ===============================================================

# Code page names from https://docs.microsoft.com/en-gb/windows/desktop/Intl/code-page-identifiers
# Retrieved on the 2018-12-13
# How it was converted to Python:
# 1) copy the table data (3 columns) from browser into Excel
# 2) use the following formula to concatenate 1st and 3rd columns: =A1 & ": " & "'" & C1 & "',"
# 3) copy from Excel into Python

CODEPAGE_NAME = {
    37: 'IBM EBCDIC US-Canada',
    437: 'OEM United States',
    500: 'IBM EBCDIC International',
    708: 'Arabic (ASMO 708)',
    709: 'Arabic (ASMO-449+, BCON V4)',
    710: 'Arabic - Transparent Arabic',
    720: 'Arabic (Transparent ASMO); Arabic (DOS)',
    737: 'OEM Greek (formerly 437G); Greek (DOS)',
    775: 'OEM Baltic; Baltic (DOS)',
    850: 'OEM Multilingual Latin 1; Western European (DOS)',
    852: 'OEM Latin 2; Central European (DOS)',
    855: 'OEM Cyrillic (primarily Russian)',
    857: 'OEM Turkish; Turkish (DOS)',
    858: 'OEM Multilingual Latin 1 + Euro symbol',
    860: 'OEM Portuguese; Portuguese (DOS)',
    861: 'OEM Icelandic; Icelandic (DOS)',
    862: 'OEM Hebrew; Hebrew (DOS)',
    863: 'OEM French Canadian; French Canadian (DOS)',
    864: 'OEM Arabic; Arabic (864)',
    865: 'OEM Nordic; Nordic (DOS)',
    866: 'OEM Russian; Cyrillic (DOS)',
    869: 'OEM Modern Greek; Greek, Modern (DOS)',
    870: 'IBM EBCDIC Multilingual/ROECE (Latin 2); IBM EBCDIC Multilingual Latin 2',
    874: 'ANSI/OEM Thai (ISO 8859-11); Thai (Windows)',
    875: 'IBM EBCDIC Greek Modern',
    932: 'ANSI/OEM Japanese; Japanese (Shift-JIS)',
    936: 'ANSI/OEM Simplified Chinese (PRC, Singapore); Chinese Simplified (GB2312)',
    949: 'ANSI/OEM Korean (Unified Hangul Code)',
    950: 'ANSI/OEM Traditional Chinese (Taiwan; Hong Kong SAR, PRC); Chinese Traditional (Big5)',
    1026: 'IBM EBCDIC Turkish (Latin 5)',
    1047: 'IBM EBCDIC Latin 1/Open System',
    1140: 'IBM EBCDIC US-Canada (037 + Euro symbol); IBM EBCDIC (US-Canada-Euro)',
    1141: 'IBM EBCDIC Germany (20273 + Euro symbol); IBM EBCDIC (Germany-Euro)',
    1142: 'IBM EBCDIC Denmark-Norway (20277 + Euro symbol); IBM EBCDIC (Denmark-Norway-Euro)',
    1143: 'IBM EBCDIC Finland-Sweden (20278 + Euro symbol); IBM EBCDIC (Finland-Sweden-Euro)',
    1144: 'IBM EBCDIC Italy (20280 + Euro symbol); IBM EBCDIC (Italy-Euro)',
    1145: 'IBM EBCDIC Latin America-Spain (20284 + Euro symbol); IBM EBCDIC (Spain-Euro)',
    1146: 'IBM EBCDIC United Kingdom (20285 + Euro symbol); IBM EBCDIC (UK-Euro)',
    1147: 'IBM EBCDIC France (20297 + Euro symbol); IBM EBCDIC (France-Euro)',
    1148: 'IBM EBCDIC International (500 + Euro symbol); IBM EBCDIC (International-Euro)',
    1149: 'IBM EBCDIC Icelandic (20871 + Euro symbol); IBM EBCDIC (Icelandic-Euro)',
    1200: 'Unicode UTF-16, little endian byte order (BMP of ISO 10646); available only to managed applications',
    1201: 'Unicode UTF-16, big endian byte order; available only to managed applications',
    1250: 'ANSI Central European; Central European (Windows)',
    1251: 'ANSI Cyrillic; Cyrillic (Windows)',
    1252: 'ANSI Latin 1; Western European (Windows)',
    1253: 'ANSI Greek; Greek (Windows)',
    1254: 'ANSI Turkish; Turkish (Windows)',
    1255: 'ANSI Hebrew; Hebrew (Windows)',
    1256: 'ANSI Arabic; Arabic (Windows)',
    1257: 'ANSI Baltic; Baltic (Windows)',
    1258: 'ANSI/OEM Vietnamese; Vietnamese (Windows)',
    1361: 'Korean (Johab)',
    10000: 'MAC Roman; Western European (Mac)',
    10001: 'Japanese (Mac)',
    10002: 'MAC Traditional Chinese (Big5); Chinese Traditional (Mac)',
    10003: 'Korean (Mac)',
    10004: 'Arabic (Mac)',
    10005: 'Hebrew (Mac)',
    10006: 'Greek (Mac)',
    10007: 'Cyrillic (Mac)',
    10008: 'MAC Simplified Chinese (GB 2312); Chinese Simplified (Mac)',
    10010: 'Romanian (Mac)',
    10017: 'Ukrainian (Mac)',
    10021: 'Thai (Mac)',
    10029: 'MAC Latin 2; Central European (Mac)',
    10079: 'Icelandic (Mac)',
    10081: 'Turkish (Mac)',
    10082: 'Croatian (Mac)',
    12000: 'Unicode UTF-32, little endian byte order; available only to managed applications',
    12001: 'Unicode UTF-32, big endian byte order; available only to managed applications',
    20000: 'CNS Taiwan; Chinese Traditional (CNS)',
    20001: 'TCA Taiwan',
    20002: 'Eten Taiwan; Chinese Traditional (Eten)',
    20003: 'IBM5550 Taiwan',
    20004: 'TeleText Taiwan',
    20005: 'Wang Taiwan',
    20105: 'IA5 (IRV International Alphabet No. 5, 7-bit); Western European (IA5)',
    20106: 'IA5 German (7-bit)',
    20107: 'IA5 Swedish (7-bit)',
    20108: 'IA5 Norwegian (7-bit)',
    20127: 'US-ASCII (7-bit)',
    20261: 'T.61',
    20269: 'ISO 6937 Non-Spacing Accent',
    20273: 'IBM EBCDIC Germany',
    20277: 'IBM EBCDIC Denmark-Norway',
    20278: 'IBM EBCDIC Finland-Sweden',
    20280: 'IBM EBCDIC Italy',
    20284: 'IBM EBCDIC Latin America-Spain',
    20285: 'IBM EBCDIC United Kingdom',
    20290: 'IBM EBCDIC Japanese Katakana Extended',
    20297: 'IBM EBCDIC France',
    20420: 'IBM EBCDIC Arabic',
    20423: 'IBM EBCDIC Greek',
    20424: 'IBM EBCDIC Hebrew',
    20833: 'IBM EBCDIC Korean Extended',
    20838: 'IBM EBCDIC Thai',
    20866: 'Russian (KOI8-R); Cyrillic (KOI8-R)',
    20871: 'IBM EBCDIC Icelandic',
    20880: 'IBM EBCDIC Cyrillic Russian',
    20905: 'IBM EBCDIC Turkish',
    20924: 'IBM EBCDIC Latin 1/Open System (1047 + Euro symbol)',
    20932: 'Japanese (JIS 0208-1990 and 0212-1990)',
    20936: 'Simplified Chinese (GB2312); Chinese Simplified (GB2312-80)',
    20949: 'Korean Wansung',
    21025: 'IBM EBCDIC Cyrillic Serbian-Bulgarian',
    21027: '(deprecated)',
    21866: 'Ukrainian (KOI8-U); Cyrillic (KOI8-U)',
    28591: 'ISO 8859-1 Latin 1; Western European (ISO)',
    28592: 'ISO 8859-2 Central European; Central European (ISO)',
    28593: 'ISO 8859-3 Latin 3',
    28594: 'ISO 8859-4 Baltic',
    28595: 'ISO 8859-5 Cyrillic',
    28596: 'ISO 8859-6 Arabic',
    28597: 'ISO 8859-7 Greek',
    28598: 'ISO 8859-8 Hebrew; Hebrew (ISO-Visual)',
    28599: 'ISO 8859-9 Turkish',
    28603: 'ISO 8859-13 Estonian',
    28605: 'ISO 8859-15 Latin 9',
    29001: 'Europa 3',
    38598: 'ISO 8859-8 Hebrew; Hebrew (ISO-Logical)',
    50220: 'ISO 2022 Japanese with no halfwidth Katakana; Japanese (JIS)',
    50221: 'ISO 2022 Japanese with halfwidth Katakana; Japanese (JIS-Allow 1 byte Kana)',
    50222: 'ISO 2022 Japanese JIS X 0201-1989; Japanese (JIS-Allow 1 byte Kana - SO/SI)',
    50225: 'ISO 2022 Korean',
    50227: 'ISO 2022 Simplified Chinese; Chinese Simplified (ISO 2022)',
    50229: 'ISO 2022 Traditional Chinese',
    50930: 'EBCDIC Japanese (Katakana) Extended',
    50931: 'EBCDIC US-Canada and Japanese',
    50933: 'EBCDIC Korean Extended and Korean',
    50935: 'EBCDIC Simplified Chinese Extended and Simplified Chinese',
    50936: 'EBCDIC Simplified Chinese',
    50937: 'EBCDIC US-Canada and Traditional Chinese',
    50939: 'EBCDIC Japanese (Latin) Extended and Japanese',
    51932: 'EUC Japanese',
    51936: 'EUC Simplified Chinese; Chinese Simplified (EUC)',
    51949: 'EUC Korean',
    51950: 'EUC Traditional Chinese',
    52936: 'HZ-GB2312 Simplified Chinese; Chinese Simplified (HZ)',
    54936: 'Windows XP and later: GB18030 Simplified Chinese (4 byte); Chinese Simplified (GB18030)',
    57002: 'ISCII Devanagari',
    57003: 'ISCII Bangla',
    57004: 'ISCII Tamil',
    57005: 'ISCII Telugu',
    57006: 'ISCII Assamese',
    57007: 'ISCII Odia',
    57008: 'ISCII Kannada',
    57009: 'ISCII Malayalam',
    57010: 'ISCII Gujarati',
    57011: 'ISCII Punjabi',
    65000: 'Unicode (UTF-7)',
    65001: 'Unicode (UTF-8)',
}


# Mapping from codepages to Python codecs, when 'cpXXX' does not work
# (inspired from http://stackoverflow.com/questions/1592925/decoding-mac-os-text-in-python)
CODEPAGE_TO_CODEC = {
    37: 'cp037',
    708: 'arabic', # not found: Arabic (ASMO 708) => arabic = iso-8859-6
    709: 'arabic', # not found: Arabic (ASMO-449+, BCON V4) => arabic = iso-8859-6
    710: 'arabic', # not found: Arabic - Transparent Arabic => arabic = iso-8859-6
    870: 'latin2', # IBM EBCDIC Multilingual/ROECE (Latin 2); IBM EBCDIC Multilingual Latin 2
    1047: 'latin1', # IBM EBCDIC Latin 1/Open System
    1141: 'cp273', # IBM EBCDIC Germany (20273 + Euro symbol); IBM EBCDIC (Germany-Euro)
    1200: 'utf_16_le', # Unicode UTF-16, little endian byte order (BMP of ISO 10646); available only to managed applications
    1201: 'utf_16_be', # Unicode UTF-16, big endian byte order; available only to managed applications

    10000: 'mac-roman',
    10001: 'shiftjis',  # not found: 'mac-shift-jis',
    10002: 'big5',      # not found: 'mac-big5',
    10003: 'ascii',     # nothing appropriate found: 'mac-hangul',
    10004: 'mac-arabic',
    10005: 'hebrew',    # not found: 'mac-hebrew',
    10006: 'mac-greek',
    #10007: 'ascii',     # nothing appropriate found: 'mac-russian',
    10007: 'mac_cyrillic',  # guess (from xlrd)
    10008: 'gb2312',    # not found: 'mac-gb2312',
    10021: 'thai',      # not found: mac-thai',
    #10029: 'maccentraleurope',  # not found: 'mac-east europe',
    10029: 'mac_latin2',  # guess (from xlrd)
    10079: 'mac_iceland',  # guess (from xlrd)
    10081: 'mac-turkish',

    12000: 'utf_32_le', # Unicode UTF-32, little endian byte order
    12001: 'utf_32_be', # Unicode UTF-32, big endian byte order

    20127: 'ascii',

    28591: 'latin1',
    28592: 'iso8859_2',
    28593: 'iso8859_3',
    28594: 'iso8859_4',
    28595: 'iso8859_5',
    28596: 'iso8859_6',
    28597: 'iso8859_7',
    28598: 'iso8859_8',
    28599: 'iso8859_9',
    28603: 'iso8859_13',
    28605: 'iso8859_15',

    32768: 'mac_roman', # from xlrd
    32769: 'cp1252', # from xlrd
    38598: 'iso8859_8',

    65000: 'utf7',
    65001: 'utf8',
}


# === FUNCTIONS ==============================================================

def codepage2codec(codepage):
    """
    convert a codepage number to a Python codec.
    If the corresponding codec cannot be found, returns "utf8" by default.

    :param codepage: int, code page number
    :return: str, Python codec name
    """
    if codepage in CODEPAGE_TO_CODEC:
        codec = CODEPAGE_TO_CODEC[codepage]
    else:
        codec = 'cp%d' % codepage
    try:
        codecs.lookup(codec)
    except LookupError:
        #log.error('Codec not found for code page %d, using UTF-8 as fallback.' % codepage)
        codec = 'utf8'
    return codec


def get_codepage_name(codepage):
    """
    return the name of a codepage based on its number
    :param codepage: int, codepage number
    :return: str, codepage name
    """
    return CODEPAGE_NAME.get(codepage, 'Unknown code page')


# === MAIN: TESTS ============================================================

if __name__ == '__main__':
    for cp in sorted(CODEPAGE_NAME.keys()):
        print('Code Page: %d => codec: %s - %s' % (cp, codepage2codec(cp), CODEPAGE_NAME[cp]))