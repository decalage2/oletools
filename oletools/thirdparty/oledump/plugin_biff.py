#!/usr/bin/env python

__description__ = 'BIFF plugin for oledump.py'
__author__ = 'Didier Stevens'
__version__ = '0.0.5'
__date__ = '2019/03/06'

# Slightly modified version by Philippe Lagadec to be imported into olevba

"""

Source code put in public domain by Didier Stevens, no Copyright
https://DidierStevens.com
Use at your own risk

History:
  2014/11/15: start
  2014/11/21: changed interface: added options; added options -a (asciidump) and -s (strings)
  2017/12/10: 0.0.2 added optparse & option -o
  2017/12/12: added option -f
  2017/12/13: added 0x support for option -f
  2018/10/24: 0.0.3 started coding Excel 4.0 macro support
  2018/10/25: continue
  2018/10/26: continue
  2019/01/05: 0.0.4 added option -x
  2019/03/06: 0.0.5 enhanced parsing of formula expressions

Todo:
"""

import struct
import re
import optparse
import binascii
import sys

# from olevba:

if sys.version_info[0] <= 2:
    # Python 2.x
    PYTHON2 = True
else:
    # Python 3.x+
    PYTHON2 = False

def unicode2str(unicode_string):
    """
    convert a unicode string to a native str:
        - on Python 3, it returns the same string
        - on Python 2, the string is encoded with UTF-8 to a bytes str
    :param unicode_string: unicode string to be converted
    :return: the string converted to str
    :rtype: str
    """
    if PYTHON2:
        return unicode_string.encode('utf8', errors='replace')
    else:
        return unicode_string


def bytes2str(bytes_string, encoding='utf8'):
    """
    convert a bytes string to a native str:
        - on Python 2, it returns the same string (bytes=str)
        - on Python 3, the string is decoded using the provided encoding
          (UTF-8 by default) to a unicode str
    :param bytes_string: bytes string to be converted
    :param encoding: codec to be used for decoding
    :return: the string converted to str
    :rtype: str
    """
    if PYTHON2:
        return bytes_string
    else:
        return bytes_string.decode(encoding, errors='replace')


dTokens = {
0x01: 'ptgExp',
0x02: 'ptgTbl',
0x03: 'ptgAdd',
0x04: 'ptgSub',
0x05: 'ptgMul',
0x06: 'ptgDiv',
0x07: 'ptgPower',
0x08: 'ptgConcat',
0x09: 'ptgLT',
0x0A: 'ptgLE',
0x0B: 'ptgEQ',
0x0C: 'ptgGE',
0x0D: 'ptgGT',
0x0E: 'ptgNE',
0x0F: 'ptgIsect',
0x10: 'ptgUnion',
0x11: 'ptgRange',
0x12: 'ptgUplus',
0x13: 'ptgUminus',
0x14: 'ptgPercent',
0x15: 'ptgParen',
0x16: 'ptgMissArg',
0x17: 'ptgStr',
0x19: 'ptgAttr',
0x1A: 'ptgSheet',
0x1B: 'ptgEndSheet',
0x1C: 'ptgErr',
0x1D: 'ptgBool',
0x1E: 'ptgInt',
0x1F: 'ptgNum',
0x20: 'ptgArray',
0x21: 'ptgFunc',
0x22: 'ptgFuncVar',
0x23: 'ptgName',
0x24: 'ptgRef',
0x25: 'ptgArea',
0x26: 'ptgMemArea',
0x27: 'ptgMemErr',
0x28: 'ptgMemNoMem',
0x29: 'ptgMemFunc',
0x2A: 'ptgRefErr',
0x2B: 'ptgAreaErr',
0x2C: 'ptgRefN',
0x2D: 'ptgAreaN',
0x2E: 'ptgMemAreaN',
0x2F: 'ptgMemNoMemN',
0x39: 'ptgNameX',
0x3A: 'ptgRef3d',
0x3B: 'ptgArea3d',
0x3C: 'ptgRefErr3d',
0x3D: 'ptgAreaErr3d',
0x40: 'ptgArrayV',
0x41: 'ptgFuncV',
0x42: 'ptgFuncVarV',
0x43: 'ptgNameV',
0x44: 'ptgRefV',
0x45: 'ptgAreaV',
0x46: 'ptgMemAreaV',
0x47: 'ptgMemErrV',
0x48: 'ptgMemNoMemV',
0x49: 'ptgMemFuncV',
0x4A: 'ptgRefErrV',
0x4B: 'ptgAreaErrV',
0x4C: 'ptgRefNV',
0x4D: 'ptgAreaNV',
0x4E: 'ptgMemAreaNV',
0x4F: 'ptgMemNoMemNV',
0x58: 'ptgFuncCEV',
0x59: 'ptgNameXV',
0x5A: 'ptgRef3dV',
0x5B: 'ptgArea3dV',
0x5C: 'ptgRefErr3dV',
0x5D: 'ptgAreaErr3dV',
0x60: 'ptgArrayA',
0x61: 'ptgFuncA',
0x62: 'ptgFuncVarA',
0x63: 'ptgNameA',
0x64: 'ptgRefA',
0x65: 'ptgAreaA',
0x66: 'ptgMemAreaA',
0x67: 'ptgMemErrA',
0x68: 'ptgMemNoMemA',
0x69: 'ptgMemFuncA',
0x6A: 'ptgRefErrA',
0x6B: 'ptgAreaErrA',
0x6C: 'ptgRefNA',
0x6D: 'ptgAreaNA',
0x6E: 'ptgMemAreaNA',
0x6F: 'ptgMemNoMemNA',
0x78: 'ptgFuncCEA',
0x79: 'ptgNameXA',
0x7A: 'ptgRef3dA',
0x7B: 'ptgArea3dA',
0x7C: 'ptgRefErr3dA',
0x7D: 'ptgAreaErr3dA',
}

#https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/00b5dd7d-51ca-4938-b7b7-483fe0e5933b
dFunctions = {
0x0000: 'COUNT',
0x0001: 'IF',
0x0002: 'ISNA',
0x0003: 'ISERROR',
0x0004: 'SUM',
0x0005: 'AVERAGE',
0x0006: 'MIN',
0x0007: 'MAX',
0x0008: 'ROW',
0x0009: 'COLUMN',
0x000A: 'NA',
0x000B: 'NPV',
0x000C: 'STDEV',
0x000D: 'DOLLAR',
0x000E: 'FIXED',
0x000F: 'SIN',
0x0010: 'COS',
0x0011: 'TAN',
0x0012: 'ATAN',
0x0013: 'PI',
0x0014: 'SQRT',
0x0015: 'EXP',
0x0016: 'LN',
0x0017: 'LOG10',
0x0018: 'ABS',
0x0019: 'INT',
0x001A: 'SIGN',
0x001B: 'ROUND',
0x001C: 'LOOKUP',
0x001D: 'INDEX',
0x001E: 'REPT',
0x001F: 'MID',
0x0020: 'LEN',
0x0021: 'VALUE',
0x0022: 'TRUE',
0x0023: 'FALSE',
0x0024: 'AND',
0x0025: 'OR',
0x0026: 'NOT',
0x0027: 'MOD',
0x0028: 'DCOUNT',
0x0029: 'DSUM',
0x002A: 'DAVERAGE',
0x002B: 'DMIN',
0x002C: 'DMAX',
0x002D: 'DSTDEV',
0x002E: 'VAR',
0x002F: 'DVAR',
0x0030: 'TEXT',
0x0031: 'LINEST',
0x0032: 'TREND',
0x0033: 'LOGEST',
0x0034: 'GROWTH',
0x0035: 'GOTO',
0x0036: 'HALT',
0x0037: 'RETURN',
0x0038: 'PV',
0x0039: 'FV',
0x003A: 'NPER',
0x003B: 'PMT',
0x003C: 'RATE',
0x003D: 'MIRR',
0x003E: 'IRR',
0x003F: 'RAND',
0x0040: 'MATCH',
0x0041: 'DATE',
0x0042: 'TIME',
0x0043: 'DAY',
0x0044: 'MONTH',
0x0045: 'YEAR',
0x0046: 'WEEKDAY',
0x0047: 'HOUR',
0x0048: 'MINUTE',
0x0049: 'SECOND',
0x004A: 'NOW',
0x004B: 'AREAS',
0x004C: 'ROWS',
0x004D: 'COLUMNS',
0x004E: 'OFFSET',
0x004F: 'ABSREF',
0x0050: 'RELREF',
0x0051: 'ARGUMENT',
0x0052: 'SEARCH',
0x0053: 'TRANSPOSE',
0x0054: 'ERROR',
0x0055: 'STEP',
0x0056: 'TYPE',
0x0057: 'ECHO',
0x0058: 'SET.NAME',
0x0059: 'CALLER',
0x005A: 'DEREF',
0x005B: 'WINDOWS',
0x005C: 'SERIES',
0x005D: 'DOCUMENTS',
0x005E: 'ACTIVE.CELL',
0x005F: 'SELECTION',
0x0060: 'RESULT',
0x0061: 'ATAN2',
0x0062: 'ASIN',
0x0063: 'ACOS',
0x0064: 'CHOOSE',
0x0065: 'HLOOKUP',
0x0066: 'VLOOKUP',
0x0067: 'LINKS',
0x0068: 'INPUT',
0x0069: 'ISREF',
0x006A: 'GET.FORMULA',
0x006B: 'GET.NAME',
0x006C: 'SET.VALUE',
0x006D: 'LOG',
0x006E: 'EXEC',
0x006F: 'CHAR',
0x0070: 'LOWER',
0x0071: 'UPPER',
0x0072: 'PROPER',
0x0073: 'LEFT',
0x0074: 'RIGHT',
0x0075: 'EXACT',
0x0076: 'TRIM',
0x0077: 'REPLACE',
0x0078: 'SUBSTITUTE',
0x0079: 'CODE',
0x007A: 'NAMES',
0x007B: 'DIRECTORY',
0x007C: 'FIND',
0x007D: 'CELL',
0x007E: 'ISERR',
0x007F: 'ISTEXT',
0x0080: 'ISNUMBER',
0x0081: 'ISBLANK',
0x0082: 'T',
0x0083: 'N',
0x0084: 'FOPEN',
0x0085: 'FCLOSE',
0x0086: 'FSIZE',
0x0087: 'FREADLN',
0x0088: 'FREAD',
0x0089: 'FWRITELN',
0x008A: 'FWRITE',
0x008B: 'FPOS',
0x008C: 'DATEVALUE',
0x008D: 'TIMEVALUE',
0x008E: 'SLN',
0x008F: 'SYD',
0x0090: 'DDB',
0x0091: 'GET.DEF',
0x0092: 'REFTEXT',
0x0093: 'TEXTREF',
0x0094: 'INDIRECT',
0x0095: 'REGISTER',
0x0096: 'CALL',
0x0097: 'ADD.BAR',
0x0098: 'ADD.MENU',
0x0099: 'ADD.COMMAND',
0x009A: 'ENABLE.COMMAND',
0x009B: 'CHECK.COMMAND',
0x009C: 'RENAME.COMMAND',
0x009D: 'SHOW.BAR',
0x009E: 'DELETE.MENU',
0x009F: 'DELETE.COMMAND',
0x00A0: 'GET.CHART.ITEM',
0x00A1: 'DIALOG.BOX',
0x00A2: 'CLEAN',
0x00A3: 'MDETERM',
0x00A4: 'MINVERSE',
0x00A5: 'MMULT',
0x00A6: 'FILES',
0x00A7: 'IPMT',
0x00A8: 'PPMT',
0x00A9: 'COUNTA',
0x00AA: 'CANCEL.KEY',
0x00AB: 'FOR',
0x00AC: 'WHILE',
0x00AD: 'BREAK',
0x00AE: 'NEXT',
0x00AF: 'INITIATE',
0x00B0: 'REQUEST',
0x00B1: 'POKE',
0x00B2: 'EXECUTE',
0x00B3: 'TERMINATE',
0x00B4: 'RESTART',
0x00B5: 'HELP',
0x00B6: 'GET.BAR',
0x00B7: 'PRODUCT',
0x00B8: 'FACT',
0x00B9: 'GET.CELL',
0x00BA: 'GET.WORKSPACE',
0x00BB: 'GET.WINDOW',
0x00BC: 'GET.DOCUMENT',
0x00BD: 'DPRODUCT',
0x00BE: 'ISNONTEXT',
0x00BF: 'GET.NOTE',
0x00C0: 'NOTE',
0x00C1: 'STDEVP',
0x00C2: 'VARP',
0x00C3: 'DSTDEVP',
0x00C4: 'DVARP',
0x00C5: 'TRUNC',
0x00C6: 'ISLOGICAL',
0x00C7: 'DCOUNTA',
0x00C8: 'DELETE.BAR',
0x00C9: 'UNREGISTER',
0x00CC: 'USDOLLAR',
0x00CD: 'FINDB',
0x00CE: 'SEARCHB',
0x00CF: 'REPLACEB',
0x00D0: 'LEFTB',
0x00D1: 'RIGHTB',
0x00D2: 'MIDB',
0x00D3: 'LENB',
0x00D4: 'ROUNDUP',
0x00D5: 'ROUNDDOWN',
0x00D6: 'ASC',
0x00D7: 'DBCS',
0x00D8: 'RANK',
0x00DB: 'ADDRESS',
0x00DC: 'DAYS360',
0x00DD: 'TODAY',
0x00DE: 'VDB',
0x00DF: 'ELSE',
0x00E0: 'ELSE.IF',
0x00E1: 'END.IF',
0x00E2: 'FOR.CELL',
0x00E3: 'MEDIAN',
0x00E4: 'SUMPRODUCT',
0x00E5: 'SINH',
0x00E6: 'COSH',
0x00E7: 'TANH',
0x00E8: 'ASINH',
0x00E9: 'ACOSH',
0x00EA: 'ATANH',
0x00EB: 'DGET',
0x00EC: 'CREATE.OBJECT',
0x00ED: 'VOLATILE',
0x00EE: 'LAST.ERROR',
0x00EF: 'CUSTOM.UNDO',
0x00F0: 'CUSTOM.REPEAT',
0x00F1: 'FORMULA.CONVERT',
0x00F2: 'GET.LINK.INFO',
0x00F3: 'TEXT.BOX',
0x00F4: 'INFO',
0x00F5: 'GROUP',
0x00F6: 'GET.OBJECT',
0x00F7: 'DB',
0x00F8: 'PAUSE',
0x00FB: 'RESUME',
0x00FC: 'FREQUENCY',
0x00FD: 'ADD.TOOLBAR',
0x00FE: 'DELETE.TOOLBAR',
0x00FF: 'User Defined Function',
0x0100: 'RESET.TOOLBAR',
0x0101: 'EVALUATE',
0x0102: 'GET.TOOLBAR',
0x0103: 'GET.TOOL',
0x0104: 'SPELLING.CHECK',
0x0105: 'ERROR.TYPE',
0x0106: 'APP.TITLE',
0x0107: 'WINDOW.TITLE',
0x0108: 'SAVE.TOOLBAR',
0x0109: 'ENABLE.TOOL',
0x010A: 'PRESS.TOOL',
0x010B: 'REGISTER.ID',
0x010C: 'GET.WORKBOOK',
0x010D: 'AVEDEV',
0x010E: 'BETADIST',
0x010F: 'GAMMALN',
0x0110: 'BETAINV',
0x0111: 'BINOMDIST',
0x0112: 'CHIDIST',
0x0113: 'CHIINV',
0x0114: 'COMBIN',
0x0115: 'CONFIDENCE',
0x0116: 'CRITBINOM',
0x0117: 'EVEN',
0x0118: 'EXPONDIST',
0x0119: 'FDIST',
0x011A: 'FINV',
0x011B: 'FISHER',
0x011C: 'FISHERINV',
0x011D: 'FLOOR',
0x011E: 'GAMMADIST',
0x011F: 'GAMMAINV',
0x0120: 'CEILING',
0x0121: 'HYPGEOMDIST',
0x0122: 'LOGNORMDIST',
0x0123: 'LOGINV',
0x0124: 'NEGBINOMDIST',
0x0125: 'NORMDIST',
0x0126: 'NORMSDIST',
0x0127: 'NORMINV',
0x0128: 'NORMSINV',
0x0129: 'STANDARDIZE',
0x012A: 'ODD',
0x012B: 'PERMUT',
0x012C: 'POISSON',
0x012D: 'TDIST',
0x012E: 'WEIBULL',
0x012F: 'SUMXMY2',
0x0130: 'SUMX2MY2',
0x0131: 'SUMX2PY2',
0x0132: 'CHITEST',
0x0133: 'CORREL',
0x0134: 'COVAR',
0x0135: 'FORECAST',
0x0136: 'FTEST',
0x0137: 'INTERCEPT',
0x0138: 'PEARSON',
0x0139: 'RSQ',
0x013A: 'STEYX',
0x013B: 'SLOPE',
0x013C: 'TTEST',
0x013D: 'PROB',
0x013E: 'DEVSQ',
0x013F: 'GEOMEAN',
0x0140: 'HARMEAN',
0x0141: 'SUMSQ',
0x0142: 'KURT',
0x0143: 'SKEW',
0x0144: 'ZTEST',
0x0145: 'LARGE',
0x0146: 'SMALL',
0x0147: 'QUARTILE',
0x0148: 'PERCENTILE',
0x0149: 'PERCENTRANK',
0x014A: 'MODE',
0x014B: 'TRIMMEAN',
0x014C: 'TINV',
0x014E: 'MOVIE.COMMAND',
0x014F: 'GET.MOVIE',
0x0150: 'CONCATENATE',
0x0151: 'POWER',
0x0152: 'PIVOT.ADD.DATA',
0x0153: 'GET.PIVOT.TABLE',
0x0154: 'GET.PIVOT.FIELD',
0x0155: 'GET.PIVOT.ITEM',
0x0156: 'RADIANS',
0x0157: 'DEGREES',
0x0158: 'SUBTOTAL',
0x0159: 'SUMIF',
0x015A: 'COUNTIF',
0x015B: 'COUNTBLANK',
0x015C: 'SCENARIO.GET',
0x015D: 'OPTIONS.LISTS.GET',
0x015E: 'ISPMT',
0x015F: 'DATEDIF',
0x0160: 'DATESTRING',
0x0161: 'NUMBERSTRING',
0x0162: 'ROMAN',
0x0163: 'OPEN.DIALOG',
0x0164: 'SAVE.DIALOG',
0x0165: 'VIEW.GET',
0x0166: 'GETPIVOTDATA',
0x0167: 'HYPERLINK',
0x0168: 'PHONETIC',
0x0169: 'AVERAGEA',
0x016A: 'MAXA',
0x016B: 'MINA',
0x016C: 'STDEVPA',
0x016D: 'VARPA',
0x016E: 'STDEVA',
0x016F: 'VARA',
0x0170: 'BAHTTEXT',
0x0171: 'THAIDAYOFWEEK',
0x0172: 'THAIDIGIT',
0x0173: 'THAIMONTHOFYEAR',
0x0174: 'THAINUMSOUND',
0x0175: 'THAINUMSTRING',
0x0176: 'THAISTRINGLENGTH',
0x0177: 'ISTHAIDIGIT',
0x0178: 'ROUNDBAHTDOWN',
0x0179: 'ROUNDBAHTUP',
0x017A: 'THAIYEAR',
0x017B: 'RTD',

0x8076: 'ALERT',
}

dOpcodes = {
    0x06: 'FORMULA : Cell Formula',
    0x0A: 'EOF : End of File',
    0x0C: 'CALCCOUNT : Iteration Count',
    0x0D: 'CALCMODE : Calculation Mode',
    0x0E: 'PRECISION : Precision',
    0x0F: 'REFMODE : Reference Mode',
    0x10: 'DELTA : Iteration Increment',
    0x11: 'ITERATION : Iteration Mode',
    0x12: 'PROTECT : Protection Flag',
    0x13: 'PASSWORD : Protection Password',
    0x14: 'HEADER : Print Header on Each Page',
    0x15: 'FOOTER : Print Footer on Each Page',
    0x16: 'EXTERNCOUNT : Number of External References',
    0x17: 'EXTERNSHEET : External Reference',
    0x18: 'LABEL : Cell Value, String Constant',
    0x19: 'WINDOWPROTECT : Windows Are Protected',
    0x1A: 'VERTICALPAGEBREAKS : Explicit Column Page Breaks',
    0x1B: 'HORIZONTALPAGEBREAKS : Explicit Row Page Breaks',
    0x1C: 'NOTE : Comment Associated with a Cell',
    0x1D: 'SELECTION : Current Selection',
    0x22: '1904 : 1904 Date System',
    0x26: 'LEFTMARGIN : Left Margin Measurement',
    0x27: 'RIGHTMARGIN : Right Margin Measurement',
    0x28: 'TOPMARGIN : Top Margin Measurement',
    0x29: 'BOTTOMMARGIN : Bottom Margin Measurement',
    0x2A: 'PRINTHEADERS : Print Row/Column Labels',
    0x2B: 'PRINTGRIDLINES : Print Gridlines Flag',
    0x2F: 'FILEPASS : File Is Password-Protected',
    0x3C: 'CONTINUE : Continues Long Records',
    0x3D: 'WINDOW1 : Window Information',
    0x40: 'BACKUP : Save Backup Version of the File',
    0x41: 'PANE : Number of Panes and Their Position',
    0x42: 'CODENAME : VBE Object Name',
    0x42: 'CODEPAGE : Default Code Page',
    0x4D: 'PLS : Environment-Specific Print Record',
    0x50: 'DCON : Data Consolidation Information',
    0x51: 'DCONREF : Data Consolidation References',
    0x52: 'DCONNAME : Data Consolidation Named References',
    0x55: 'DEFCOLWIDTH : Default Width for Columns',
    0x59: 'XCT : CRN Record Count',
    0x5A: 'CRN : Nonresident Operands',
    0x5B: 'FILESHARING : File-Sharing Information',
    0x5C: 'WRITEACCESS : Write Access User Name',
    0x5D: 'OBJ : Describes a Graphic Object',
    0x5E: 'UNCALCED : Recalculation Status',
    0x5F: 'SAVERECALC : Recalculate Before Save',
    0x60: 'TEMPLATE : Workbook Is a Template',
    0x63: 'OBJPROTECT : Objects Are Protected',
    0x7D: 'COLINFO : Column Formatting Information',
    0x7E: 'RK : Cell Value, RK Number',
    0x7F: 'IMDATA : Image Data',
    0x80: 'GUTS : Size of Row and Column Gutters',
    0x81: 'WSBOOL : Additional Workspace Information',
    0x82: 'GRIDSET : State Change of Gridlines Option',
    0x83: 'HCENTER : Center Between Horizontal Margins',
    0x84: 'VCENTER : Center Between Vertical Margins',
    0x85: 'BOUNDSHEET : Sheet Information',
    0x86: 'WRITEPROT : Workbook Is Write-Protected',
    0x87: 'ADDIN : Workbook Is an Add-in Macro',
    0x88: 'EDG : Edition Globals',
    0x89: 'PUB : Publisher',
    0x8C: 'COUNTRY : Default Country and WIN.INI Country',
    0x8D: 'HIDEOBJ : Object Display Options',
    0x90: 'SORT : Sorting Options',
    0x91: 'SUB : Subscriber',
    0x92: 'PALETTE : Color Palette Definition',
    0x94: 'LHRECORD : .WK? File Conversion Information',
    0x95: 'LHNGRAPH : Named Graph Information',
    0x96: 'SOUND : Sound Note',
    0x98: 'LPR : Sheet Was Printed Using LINE.PRINT(',
    0x99: 'STANDARDWIDTH : Standard Column Width',
    0x9A: 'FNGROUPNAME : Function Group Name',
    0x9B: 'FILTERMODE : Sheet Contains Filtered List',
    0x9C: 'FNGROUPCOUNT : Built-in Function Group Count',
    0x9D: 'AUTOFILTERINFO : Drop-Down Arrow Count',
    0x9E: 'AUTOFILTER : AutoFilter Data',
    0xA0: 'SCL : Window Zoom Magnification',
    0xA1: 'SETUP : Page Setup',
    0xA9: 'COORDLIST : Polygon Object Vertex Coordinates',
    0xAB: 'GCW : Global Column-Width Flags',
    0xAE: 'SCENMAN : Scenario Output Data',
    0xAF: 'SCENARIO : Scenario Data',
    0xB0: 'SXVIEW : View Definition',
    0xB1: 'SXVD : View Fields',
    0xB2: 'SXVI : View Item',
    0xB4: 'SXIVD : Row/Column Field IDs',
    0xB5: 'SXLI : Line Item Array',
    0xB6: 'SXPI : Page Item',
    0xB8: 'DOCROUTE : Routing Slip Information',
    0xB9: 'RECIPNAME : Recipient Name',
    0xBC: 'SHRFMLA : Shared Formula',
    0xBD: 'MULRK : Multiple  RK Cells',
    0xBE: 'MULBLANK : Multiple Blank Cells',
    0xC1: 'MMS :  ADDMENU / DELMENU Record Group Count',
    0xC2: 'ADDMENU : Menu Addition',
    0xC3: 'DELMENU : Menu Deletion',
    0xC5: 'SXDI : Data Item',
    0xC6: 'SXDB : PivotTable Cache Data',
    0xCD: 'SXSTRING : String',
    0xD0: 'SXTBL : Multiple Consolidation Source Info',
    0xD1: 'SXTBRGIITM : Page Item Name Count',
    0xD2: 'SXTBPG : Page Item Indexes',
    0xD3: 'OBPROJ : Visual Basic Project',
    0xD5: 'SXIDSTM : Stream ID',
    0xD6: 'RSTRING : Cell with Character Formatting',
    0xD7: 'DBCELL : Stream Offsets',
    0xDA: 'BOOKBOOL : Workbook Option Flag',
    0xDC: 'PARAMQRY : Query Parameters',
    0xDC: 'SXEXT : External Source Information',
    0xDD: 'SCENPROTECT : Scenario Protection',
    0xDE: 'OLESIZE : Size of OLE Object',
    0xDF: 'UDDESC : Description String for Chart Autoformat',
    0xE0: 'XF : Extended Format',
    0xE1: 'INTERFACEHDR : Beginning of User Interface Records',
    0xE2: 'INTERFACEEND : End of User Interface Records',
    0xE3: 'SXVS : View Source',
    0xE5: 'MERGECELLS : Merged Cells',
    0xEA: 'TABIDCONF : Sheet Tab ID of Conflict History',
    0xEB: 'MSODRAWINGGROUP : Microsoft Office Drawing Group',
    0xEC: 'MSODRAWING : Microsoft Office Drawing',
    0xED: 'MSODRAWINGSELECTION : Microsoft Office Drawing Selection',
    0xF0: 'SXRULE : PivotTable Rule Data',
    0xF1: 'SXEX : PivotTable View Extended Information',
    0xF2: 'SXFILT : PivotTable Rule Filter',
    0xF4: 'SXDXF : Pivot Table Formatting',
    0xF5: 'SXITM : Pivot Table Item Indexes',
    0xF6: 'SXNAME : PivotTable Name',
    0xF7: 'SXSELECT : PivotTable Selection Information',
    0xF8: 'SXPAIR : PivotTable Name Pair',
    0xF9: 'SXFMLA : Pivot Table Parsed Expression',
    0xFB: 'SXFORMAT : PivotTable Format Record',
    0xFC: 'SST : Shared String Table',
    0xFD: 'LABELSST : Cell Value, String Constant/ SST',
    0xFF: 'EXTSST : Extended Shared String Table',
    0x100: 'SXVDEX : Extended PivotTable View Fields',
    0x103: 'SXFORMULA : PivotTable Formula Record',
    0x122: 'SXDBEX : PivotTable Cache Data',
    0x13D: 'TABID : Sheet Tab Index Array',
    0x160: 'USESELFS : Natural Language Formulas Flag',
    0x161: 'DSF : Double Stream File',
    0x162: 'XL5MODIFY : Flag for  DSF',
    0x1A5: 'FILESHARING2 : File-Sharing Information for Shared Lists',
    0x1A9: 'USERBVIEW : Workbook Custom View Settings',
    0x1AA: 'USERSVIEWBEGIN : Custom View Settings',
    0x1AB: 'USERSVIEWEND : End of Custom View Records',
    0x1AD: 'QSI : External Data Range',
    0x1AE: 'SUPBOOK : Supporting Workbook',
    0x1AF: 'PROT4REV : Shared Workbook Protection Flag',
    0x1B0: 'CONDFMT : Conditional Formatting Range Information',
    0x1B1: 'CF : Conditional Formatting Conditions',
    0x1B2: 'DVAL : Data Validation Information',
    0x1B5: 'DCONBIN : Data Consolidation Information',
    0x1B6: 'TXO : Text Object',
    0x1B7: 'REFRESHALL : Refresh Flag',
    0x1B8: 'HLINK : Hyperlink',
    0x1BB: 'SXFDBTYPE : SQL Datatype Identifier',
    0x1BC: 'PROT4REVPASS : Shared Workbook Protection Password',
    0x1BE: 'DV : Data Validation Criteria',
    0x1C0: 'EXCEL9FILE : Excel 9 File',
    0x1C1: 'RECALCID : Recalc Information',
    0x200: 'DIMENSIONS : Cell Table Size',
    0x201: 'BLANK : Cell Value, Blank Cell',
    0x203: 'NUMBER : Cell Value, Floating-Point Number',
    0x204: 'LABEL : Cell Value, String Constant',
    0x205: 'BOOLERR : Cell Value, Boolean or Error',
    0x207: 'STRING : String Value of a Formula',
    0x208: 'ROW : Describes a Row',
    0x20B: 'INDEX : Index Record',
    0x218: 'NAME : Defined Name',
    0x221: 'ARRAY : Array-Entered Formula',
    0x223: 'EXTERNNAME : Externally Referenced Name',
    0x225: 'DEFAULTROWHEIGHT : Default Row Height',
    0x231: 'FONT : Font Description',
    0x236: 'TABLE : Data Table',
    0x23E: 'WINDOW2 : Sheet Window Information',
    0x293: 'STYLE : Style Information',
    0x406: 'FORMULA : Cell Formula',
    0x41E: 'FORMAT : Number Format',
    0x800: 'HLINKTOOLTIP : Hyperlink Tooltip',
    0x801: 'WEBPUB : Web Publish Item',
    0x802: 'QSISXTAG : PivotTable and Query Table Extensions',
    0x803: 'DBQUERYEXT : Database Query Extensions',
    0x804: 'EXTSTRING :  FRT String',
    0x805: 'TXTQUERY : Text Query Information',
    0x806: 'QSIR : Query Table Formatting',
    0x807: 'QSIF : Query Table Field Formatting',
    0x809: 'BOF : Beginning of File',
    0x80A: 'OLEDBCONN : OLE Database Connection',
    0x80B: 'WOPT : Web Options',
    0x80C: 'SXVIEWEX : Pivot Table OLAP Extensions',
    0x80D: 'SXTH : PivotTable OLAP Hierarchy',
    0x80E: 'SXPIEX : OLAP Page Item Extensions',
    0x80F: 'SXVDTEX : View Dimension OLAP Extensions',
    0x810: 'SXVIEWEX9 : Pivot Table Extensions',
    0x812: 'CONTINUEFRT : Continued  FRT',
    0x813: 'REALTIMEDATA : Real-Time Data (RTD)',
    0x862: 'SHEETEXT : Extra Sheet Info',
    0x863: 'BOOKEXT : Extra Book Info',
    0x864: 'SXADDL : Pivot Table Additional Info',
    0x865: 'CRASHRECERR : Crash Recovery Error',
    0x866: 'HFPicture : Header / Footer Picture',
    0x867: 'FEATHEADR : Shared Feature Header',
    0x868: 'FEAT : Shared Feature Record',
    0x86A: 'DATALABEXT : Chart Data Label Extension',
    0x86B: 'DATALABEXTCONTENTS : Chart Data Label Extension Contents',
    0x86C: 'CELLWATCH : Cell Watch',
    0x86d: 'FEATINFO : Shared Feature Info Record',
    0x871: 'FEATHEADR11 : Shared Feature Header 11',
    0x872: 'FEAT11 : Shared Feature 11 Record',
    0x873: 'FEATINFO11 : Shared Feature Info 11 Record',
    0x874: 'DROPDOWNOBJIDS : Drop Down Object',
    0x875: 'CONTINUEFRT11 : Continue  FRT 11',
    0x876: 'DCONN : Data Connection',
    0x877: 'LIST12 : Extra Table Data Introduced in Excel 2007',
    0x878: 'FEAT12 : Shared Feature 12 Record',
    0x879: 'CONDFMT12 : Conditional Formatting Range Information 12',
    0x87A: 'CF12 : Conditional Formatting Condition 12',
    0x87B: 'CFEX : Conditional Formatting Extension',
    0x87C: 'XFCRC : XF Extensions Checksum',
    0x87D: 'XFEXT : XF Extension',
    0x87E: 'EZFILTER12 : AutoFilter Data Introduced in Excel 2007',
    0x87F: 'CONTINUEFRT12 : Continue FRT 12',
    0x881: 'SXADDL12 : Additional Workbook Connections Information',
    0x884: 'MDTINFO : Information about a Metadata Type',
    0x885: 'MDXSTR : MDX Metadata String',
    0x886: 'MDXTUPLE : Tuple MDX Metadata',
    0x887: 'MDXSET : Set MDX Metadata',
    0x888: 'MDXPROP : Member Property MDX Metadata',
    0x889: 'MDXKPI : Key Performance Indicator MDX Metadata',
    0x88A: 'MDTB : Block of Metadata Records',
    0x88B: 'PLV : Page Layout View Settings in Excel 2007',
    0x88C: 'COMPAT12 : Compatibility Checker 12',
    0x88D: 'DXF : Differential XF',
    0x88E: 'TABLESTYLES : Table Styles',
    0x88F: 'TABLESTYLE : Table Style',
    0x890: 'TABLESTYLEELEMENT : Table Style Element',
    0x892: 'STYLEEXT : Named Cell Style Extension',
    0x893: 'NAMEPUBLISH : Publish To Excel Server Data for Name',
    0x894: 'NAMECMT : Name Comment',
    0x895: 'SORTDATA12 : Sort Data 12',
    0x896: 'THEME : Theme',
    0x897: 'GUIDTYPELIB : VB Project Typelib GUID',
    0x898: 'FNGRP12 : Function Group',
    0x899: 'NAMEFNGRP12 : Extra Function Group',
    0x89A: 'MTRSETTINGS : Multi-Threaded Calculation Settings',
    0x89B: 'COMPRESSPICTURES : Automatic Picture Compression Mode',
    0x89C: 'HEADERFOOTER : Header Footer',
    0x8A3: 'FORCEFULLCALCULATION : Force Full Calculation Settings',
    0x8c1: 'LISTOBJ : List Object',
    0x8c2: 'LISTFIELD : List Field',
    0x8c3: 'LISTDV : List Data Validation',
    0x8c4: 'LISTCONDFMT : List Conditional Formatting',
    0x8c5: 'LISTCF : List Cell Formatting',
    0x8c6: 'FMQRY : Filemaker queries',
    0x8c7: 'FMSQRY : File maker queries',
    0x8c8: 'PLV : Page Layout View in Mac Excel 11',
    0x8c9: 'LNEXT : Extension information for borders in Mac Office 11',
    0x8ca: 'MKREXT : Extension information for markers in Mac Office 11'
}


# CIC: Call If Callable
def CIC(expression):
    if callable(expression):
        return expression()
    else:
        return expression


# IFF: IF Function
def IFF(expression, valueTrue, valueFalse):
    if expression:
        return CIC(valueTrue)
    else:
        return CIC(valueFalse)


def CombineHexASCII(hexDump, asciiDump, length):
    if hexDump == '':
        return ''
    return hexDump + '  ' + (' ' * (3 * (length - len(asciiDump)))) + asciiDump

def HexASCII(data, length=16):
    result = []
    if len(data) > 0:
        hexDump = ''
        asciiDump = ''
        for i, b in enumerate(data):
            if i % length == 0:
                if hexDump != '':
                    result.append(CombineHexASCII(hexDump, asciiDump, length))
                hexDump = '%08X:' % i
                asciiDump = ''
            hexDump += ' %02X' % ord(b)
            asciiDump += IFF(ord(b) >= 32, b, '.')
        result.append(CombineHexASCII(hexDump, asciiDump, length))
    return result

def StringsASCII(data):
    """
    Extract a list of plain ASCII strings of 4+ chars found in data.
    :param data: bytearray or bytes
    :return: list of str (converted to unicode on Python 3)
    """
    # list of bytes strings:
    bytes_strings = re.findall(b'[^\x00-\x08\x0A-\x1F\x7F-\xFF]{4,}', bytes(data))
    return [bytes2str(bs) for bs in bytes_strings]

def StringsUNICODE(data):
    """
    Extract a list of Unicode strings (made of 4+ plain ASCII characters only) found in data.
    :param data: bytearray or bytes
    :return: list of str (converted to unicode on Python 3)
    """
    # list of bytes strings:
    # TODO: check if the null byte should be before or after the ascii byte
    bytes_strings = [foundunicodestring.replace(b'\x00', b'') for foundunicodestring, dummy in re.findall(b'(([^\x00-\x08\x0A-\x1F\x7F-\xFF]\x00){4,})', bytes(data))]
    return [bytes2str(bs) for bs in bytes_strings]

def Strings(data, encodings='sL'):
    """

    :param data bytearray: bytearray, data to be scanned for strings
    :param encodings:
    :return: dict with key = 's' or 'L', values = list of str
    """
    dStrings = {}
    for encoding in encodings:
        if encoding == 's':
            dStrings[encoding] = StringsASCII(data)
        elif encoding == 'L':
            dStrings[encoding] = StringsUNICODE(data)
    return dStrings

def ContainsWord(word, expression):
    return struct.pack('<H', word) in expression

# https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/6e5eed10-5b77-43d6-8dd0-37345f8654ad
def ParseLoc(expression):
    """

    :param expression bytearray: bytearray, data to be parsed
    :return:
    :rtype: str
    """
    formatcodes = 'HH'
    formatsize = struct.calcsize(formatcodes)
    row, column = struct.unpack(formatcodes, expression[0:formatsize])
    rowRelative = column & 0x8000
    colRelative = column & 0x4000
    column = column & 0x3FFF
    if rowRelative:
        rowindicator = '~'
    else:
        rowindicator = ''
        row += 1
    if colRelative:
        colindicator = '~'
    else:
        colindicator = ''
        column += 1
    return 'R%s%dC%s%d' % (rowindicator, row, colindicator, column)

def ParseExpression(expression):
    '''
    Parse an expression into a human readable string.

    :param expression bytearray: bytearray, expression data to be parsed
    :return: str, parsed expression as a string (bytes on Python 2, unicode on python 3)
    :rtype: str
    '''
    result = ''
    while len(expression) > 0:
        ptgid = expression[0]  # int
        expression = expression[1:]  # bytearray
        if ptgid in dTokens:
            result += dTokens[ptgid] + ' '
            if ptgid == 0x17:  # ptgStr
                length = expression[0]  # int
                expression = expression[1:]
                if expression[0] == 0:  # probably BIFF8 -> UNICODE (compressed)
                    expression = expression[1:]
                result += '"%s" ' % bytes2str(expression[:length])
                expression = expression[length:]
            elif ptgid == 0x19:  # ptgAttr
                grbit = expression[0]  # int
                expression = expression[1:]
                if grbit & 0x04:
                    result += 'CHOOSE '
                    break
                else:
                    expression = expression[2:]
            elif ptgid == 0x16 or ptgid == 0x0e:  # 0x0E: 'ptgNE', 0x16: 'ptgMissArg'
                pass
            elif ptgid == 0x1e:  # ptgInt
                result += '%d ' % (expression[0] + expression[1] * 0x100)
                expression = expression[2:]
            elif ptgid == 0x41:  # ptgFuncV
                functionid = expression[0] + expression[1] * 0x100
                result += '%s (0x%04x) ' % (dFunctions.get(functionid, '*UNKNOWN FUNCTION*'), functionid)
                expression = expression[2:]
            elif ptgid == 0x22 or ptgid == 0x42:  # 0x22: 'ptgFuncVar', 0x42: 'ptgFuncVarV'
                functionid = expression[1] + expression[2] * 0x100
                result += 'args %d func %s (0x%04x) ' % (expression[0], dFunctions.get(functionid, '*UNKNOWN FUNCTION*'), functionid)
                expression = expression[3:]
            elif ptgid == 0x23:  # ptgName
                result += '%04x ' % (expression[0] + expression[1] * 0x100)
                # TODO: looks like we're skipping quite a few bytes
                expression = expression[14:]
            elif ptgid == 0x1f:  # ptgNum
                result += 'FLOAT '
                # TODO: looks like we're skipping quite a few bytes
                expression = expression[8:]
            elif ptgid == 0x26:  # ptgMemArea
                expression = expression[4:]  # skipping 4 bytes
                expression = expression[expression[0] + expression[1] * 0x100:]
                result += 'REFERENCE-EXPRESSION '
            elif ptgid == 0x01:  # ptgExp
                formatcodes = 'HH'
                formatsize = struct.calcsize(formatcodes)
                row, column = struct.unpack(formatcodes, expression[0:formatsize])
                expression = expression[formatsize:]
                result += 'R%dC%d ' % (row + 1, column + 1)
            elif ptgid == 0x24 or ptgid == 0x44:  # 0x24: 'ptgRef', 0x44: 'ptgRefV'
                result += '%s ' % ParseLoc(expression)
                expression = expression[4:]
            elif ptgid == 0x3A or ptgid == 0x5A:  # 0x3A: 'ptgRef3d', 0x5A: 'ptgRef3dV'
                result += '%s ' % ParseLoc(expression[2:])
                expression = expression[6:]
            else:
                break
        else:
            result += '*UNKNOWN TOKEN* '
            break
    if len(expression) == 0:
        return result
    else:
        # 0x006E: 'EXEC', 0x0095: 'REGISTER'
        functions = [dFunctions[functionid] for functionid in [0x6E, 0x95] if ContainsWord(functionid, expression)]
        if functions != []:
            message = ' Could contain following functions: ' + ','.join(functions) + ' -'
        else:
            message = ''
        return result + ' *INCOMPLETE FORMULA PARSING*' + message + ' Remaining, unparsed expression: ' + repr(expression)
        

class cBIFF(object):  # cPluginParent):
    macroOnly = False
    name = 'BIFF plugin'

    def __init__(self, name, stream, options):
        self.streamname = name
        self.stream = stream
        self.options = options
        self.ran = False

    def Analyze(self):
        result = []
        macros4Found = False
        if self.streamname in [['Workbook'], ['Book']]:
            self.ran = True
            # use a bytearray to have Python 2+3 compatibility with the same code (no need for ord())
            stream = bytearray(self.stream)

            oParser = optparse.OptionParser()
            oParser.add_option('-s', '--strings', action='store_true', default=False, help='Dump strings')
            oParser.add_option('-a', '--hexascii', action='store_true', default=False, help='Dump hex ascii')
            oParser.add_option('-x', '--xlm', action='store_true', default=False, help='Select all records relevant for Excel 4.0 macros')
            oParser.add_option('-o', '--opcode', type=str, default='', help='Opcode to filter for')
            oParser.add_option('-f', '--find', type=str, default='', help='Content to search for')
            (options, args) = oParser.parse_args(self.options.split(' '))

            if options.find.startswith('0x'):
                options.find = binascii.a2b_hex(options.find[2:])

            while len(stream)>0:
                formatcodes = 'HH'
                formatsize = struct.calcsize(formatcodes)
                # print('formatsize=%d' % formatsize)
                opcode, length = struct.unpack(formatcodes, stream[0:formatsize])
                # print('opcode=%d length=%d len(stream)=%d' % (opcode, length, len(stream)))
                stream = stream[formatsize:]
                data = stream[:length]
                stream = stream[length:]

                if opcode in dOpcodes:
                    opcodename = dOpcodes[opcode]
                else:
                    opcodename = ''
                line = '%04x %6d %s' % (opcode, length, opcodename)
                # print(line)

                # FORMULA record
                if opcode == 0x06 and len(data) >= 21:
                    formatcodes = 'HH'
                    formatsize = struct.calcsize(formatcodes)
                    row, column = struct.unpack(formatcodes, data[0:formatsize])
                    formatcodes = 'H'
                    formatsize = struct.calcsize(formatcodes)
                    length = struct.unpack(formatcodes, data[20:20 + formatsize])[0]
                    expression = data[22:]
                    line += ' - R%dC%d len=%d %s' % (row + 1, column + 1, length, ParseExpression(expression))
                    # print(line)

                # FORMULA record #a# difference BIFF4 and BIFF5+
                if opcode == 0x18 and len(data) >= 16:
                    if data[0] & 0x20:
                        dBuildInNames = {1: 'Auto_Open', 2: 'Auto_Close'}
                        code = data[14]
                        if code == 0: #a# hack with BIFF8 Unicode
                            code = data[15]
                        line += ' - build-in-name %d %s' % (code, dBuildInNames.get(code, '?'))
                    else:
                        pass
                        line += ' - %s' % bytes2str(data[14:14+data[3]])
                    # print(line)

                # BOUNDSHEET record
                if opcode == 0x85 and len(data) >= 6:
                    dSheetType = {0: 'worksheet or dialog sheet', 1: 'Excel 4.0 macro sheet', 2: 'chart', 6: 'Visual Basic module'}
                    if data[5] == 1:
                        macros4Found = True
                    dSheetState = {0: 'visible', 1: 'hidden', 2: 'very hidden'}
                    line += ' - %s, %s' % (dSheetType.get(data[5], '%02x' % data[5]), dSheetState.get(data[4], '%02x' % data[4]))
                    # print(line)

                # STRING record
                if opcode == 0x207 and len(data) >= 4:
                    values = list(Strings(data[3:]).values())
                    strings = ''
                    if values[0] != []:
                        strings += ' '.join(values[0])
                    if values[1] != []:
                        if strings != '':
                            strings += ' '
                        strings += ' '.join(values[1])
                    line += ' - %s' % strings
                    # print(line)

                if options.find == '' and options.opcode == '' and not options.xlm or options.opcode != '' and options.opcode.lower() in line.lower() or options.find != '' and options.find in data or options.xlm and opcode in [0x06, 0x18, 0x85, 0x207]:
                    result.append(line)

                    if options.hexascii:
                        result.extend(' ' + foundstring for foundstring in HexASCII(data, 8))
                    elif options.strings:
                        dEncodings = {'s': 'ASCII', 'L': 'UNICODE'}
                        for encoding, strings in Strings(data).items():
                            if len(strings) > 0:
                                result.append(' ' + dEncodings[encoding] + ':')
                                result.extend('  ' + foundstring for foundstring in strings)

            if options.xlm and not macros4Found:
                result = []

        return result

# AddPlugin(cBIFF)
