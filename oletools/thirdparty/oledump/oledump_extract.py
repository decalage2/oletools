#!/usr/bin/env python

# Small extract of oledump.py to be able to run plugin_biff from olevba

__description__ = 'Analyze OLE files (Compound Binary Files)'
__author__ = 'Didier Stevens'
__version__ = '0.0.49'
__date__ = '2020/03/28'

"""

Source code put in public domain by Didier Stevens, no Copyright
https://DidierStevens.com
Use at your own risk
"""

class cPluginParent():
    macroOnly = False
    indexQuiet = False

plugins = []

def AddPlugin(cClass):
    global plugins

    plugins.append(cClass)


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

def P23Ord(value):
    if type(value) == int:
        return value
    else:
        return ord(value)

def P23Chr(value):
    if type(value) == int:
        return chr(value)
    else:
        return value
