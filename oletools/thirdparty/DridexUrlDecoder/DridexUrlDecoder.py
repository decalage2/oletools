# Written by @JamesHabben
# https://github.com/JamesHabben/MalwareStuff

# 2015-01-27 Slight modifications from Philippe Lagadec (PL) to use it from olevba

import sys

def DridexUrlDecode (inputText) :
    work = inputText[4:-4]
    strKeyEnc = StripCharsWithZero(work[(len(work) / 2) - 2: (len(work) / 2)])
    strKeySize = StripCharsWithZero(work[(len(work) / 2): (len(work) / 2) + 2])
    nCharSize = strKeySize - strKeyEnc
    work = work[:(len(work) / 2) - 2] + work[(len(work) / 2) + 2:]
    strKeyEnc2 = StripChars(work[(len(work) / 2) - (nCharSize/2): (len(work) / 2) + (nCharSize/2)])
    work = work[:(len(work) / 2) - (nCharSize/2)] + work[(len(work) / 2) + (nCharSize/2):]
    work_split = [work[i:i+nCharSize] for i in range(0, len(work), nCharSize)]
    decoded = ''
    for group in work_split:
        # sys.stdout.write(chr(StripChars(group)/strKeyEnc2))
        decoded += chr(StripChars(group)/strKeyEnc2)
    return decoded

def StripChars (input) :
    result = ''
    for c in input :
        if c.isdigit() :
            result += c
    return int(result)

def StripCharsWithZero (input) :
    result = ''
    for c in input :
        if c.isdigit() :
            result += c
        else:
            result += '0'
    return int(result)


# DridexUrlDecode("C3iY1epSRGe6q8g15xStVesdG717MAlg2H4hmV1vkL6Glnf0cknj")
# DridexUrlDecode("HLIY3Nf3z2k8jD37h1n2OM3N712DGQ3c5M841RZ8C5e6P1C50C4ym1oF504WyV182p4mJ16cK9Z61l47h2dU1rVB5V681sFY728i16H3E2Qm1fn47y2cgAo156j8T1s600hukKO1568X1xE4Z7d2q17jvcwgk816Yz32o9Q216Mpr0B01vcwg856a17b9j2zAmWf1536B1t7d92rI1FZ5E36Pu1jl504Z34tm2R43i55Lg2F3eLE3T28lLX1D504348Goe8Gbdp37w443ADy36X0h14g7Wb2G3u584kEG332Ut8ws3wO584pzSTf")
# DridexUrlDecode("YNPH1W47E211z3P6142cM4115K2J1696CURf1712N1OCJwc0w6Z16840Z1r600W16Z3273k6SR16Bf161Q92a016Vr16V1pc")
