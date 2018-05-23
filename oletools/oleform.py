#!/usr/bin/env python

# CHANGELOG:
# 2018-02-19 v0.53 PL: - fixed issue #260, removed long integer literals

import struct

class OleFormParsingError(Exception):
    pass

class Mask(object):
    def __init__(self, val):
        self._val = [(val & (1<<i))>>i for i in range(self._size)]

    def __str__(self):
        return ', '.join(self._names[i] for i in range(self._size) if self._val[i])

    def __getattr__(self, name):
        return self._val[self._names.index(name)]

    def __len__(self):
        return self.size

    def __getitem__(self, key):
        return self._val[self._names.index(key)]

class FormPropMask(Mask):
    """FormPropMask: [MS-OFORMS] 2.2.10.2"""
    _size = 28
    _names = ['Unused1', 'fBackColor', 'fForeColor', 'fNextAvailableID', 'Unused2_0', 'Unused2_1',
              'fBooleanProperties', 'fBooleanProperties', 'fMousePointer', 'fScrollBars',
              'fDisplayedSize', 'fLogicalSize', 'fScrollPosition', 'fGroupCnt', 'Reserved',
              'fMouseIcon', 'fCycle', 'fSpecialEffect', 'fBorderColor', 'fCaption', 'fFont',
              'fPicture', 'fZoom', 'fPictureAlignment', 'fPictureTiling', 'fPictureSizeMode',
              'fShapeCookie', 'fDrawBuffer']

class SitePropMask(Mask):
    """SitePropMask: [MS-OFORMS] 2.2.10.12.2"""
    _size = 15
    _names = ['fName', 'fTag', 'fID', 'fHelpContextID', 'fBitFlags', 'fObjectStreamSize',
              'fTabIndex', 'fClsidCacheIndex', 'fPosition', 'fGroupID', 'Unused1',
              'fControlTipText', 'fRuntimeLicKey', 'fControlSource', 'fRowSource']

class MorphDataPropMask(Mask):
    """MorphDataPropMask: [MS-OFORMS] 2.2.5.2"""
    _size = 33
    _names = ['fVariousPropertyBits', 'fBackColor', 'fForeColor', 'fMaxLength', 'fBorderStyle',
              'fScrollBars', 'fDisplayStyle', 'fMousePointer', 'fSize', 'fPasswordChar',
              'fListWidth', 'fBoundColumn', 'fTextColumn', 'fColumnCount', 'fListRows',
              'fcColumnInfo', 'fMatchEntry', 'fListStyle', 'fShowDropButtonWhen', 'UnusedBits1',
              'fDropButtonStyle', 'fMultiSelect', 'fValue', 'fCaption', 'fPicturePosition',
              'fBorderColor', 'fSpecialEffect', 'fMouseIcon', 'fPicture', 'fAccelerator',
              'UnusedBits2', 'Reserved', 'fGroupName']

class ImagePropMask(Mask):
    """ImagePropMask: [MS-OFORMS] 2.2.3.2"""
    _size = 15
    _names = ['UnusedBits1_1', 'UnusedBits1_2', 'fAutoSize', 'fBorderColor', 'fBackColor',
              'fBorderStyle', 'fMousePointer', 'fPictureSizeMode', 'fSpecialEffect', 'fSize',
              'fPicture', 'fPictureAlignment', 'fPictureTiling', 'fVariousPropertyBits',
              'fMouseIcon']

class CommandButtonPropMask(Mask):
    """CommandButtonPropMask: [MS-OFORMS] 2.2.1.2"""
    _size = 11
    _names = ['fForeColor', 'fBackColor', 'fVariousPropertyBits', 'fCaption', 'fPicturePosition',
              'fSize', 'fMousePointer', 'fPicture', 'fAccelerator', 'fTakeFocusOnClick',
              'fMouseIcon']

class SpinButtonPropMask(Mask):
    """SpinButtonPropMask: [MS-OFORMS] 2.2.8.2"""
    _size = 15
    _names = ['fForeColor', 'fBackColor', 'fVariousPropertyBits', 'fSize', 'UnusedBits1',
              'fMin', 'fMax', 'fPosition', 'fPrevEnabled', 'fNextEnabled', 'fSmallChange',
              'fOrientation', 'fDelay', 'fMouseIcon', 'fMousePointer']

class TabStripPropMask(Mask):
    """TabStripPropMask: [MS-OFORMS] 2.2.9.2"""
    _size = 25
    _names = ['fListIndex', 'fBackColor', 'fForeColor', 'Unused1', 'fSize', 'fItems',
              'fMousePointer', 'Unused2', 'fTabOrientation', 'fTabStyle', 'fMultiRow',
              'fTabFixedWidth', 'fTabFixedHeight', 'fTooltips', 'Unused3', 'fTipStrings',
              'Unused4', 'fNames', 'fVariousPropertyBits', 'fNewVersion', 'fTabsAllocated',
              'fTags', 'fTabData', 'fAccelerator', 'fMouseIcon']

class LabelPropMask(Mask):
    """LabelPropMask: [MS-OFORMS] 2.2.4.2"""
    _size = 13
    _names = ['fForeColor', 'fBackColor', 'fVariousPropertyBits', 'fCaption',
              'fPicturePosition', 'fSize', 'fMousePointer', 'fBorderColor', 'fBorderStyle',
              'fSpecialEffect', 'fPicture', 'fAccelerator', 'fMouseIcon']

class ScrollBarPropMask(Mask):
    """ScrollBarPropMask: [MS-OFORMS] 2.2.7.2"""
    _size = 17
    _names = ['fForeColor', 'fBackColor', 'fVariousPropertyBits', 'fSize', 'fMousePointer',
              'fMin', 'fMax', 'fPosition', 'UnusedBits1', 'fPrevEnabled', 'fNextEnabled',
              'fSmallChange', 'fLargeChange', 'fOrientation', 'fProportionalThumb',
              'fDelay', 'fMouseIcon']

class ExtendedStream(object):
    def __init__(self, stream, path):
        self._pos = 0
        self._jumps = []
        self._stream = stream
        self._path = path

    @classmethod
    def open(cls, ole_file, path):
        # import oletools.thirdparty.olefile as olefile
        # olefile.enable_logging()
        stream = ole_file.openstream(path)
        # print('Opening OLE stream %r - size: %d' % (path, stream.size))
        # print('declared size: %d' % ole_file.get_size(path))
        return cls(stream, path)

    def read(self, size):
        self._pos += size
        return self._stream.read(size)

    def will_jump_to(self, size):
        self._next_jump = (True, size)
        return self

    def will_pad(self, pad=4):
        self._next_jump = (False, pad)
        return self

    def __enter__(self):
        (jump_type, size) = self._next_jump
        self._jumps.append((self._pos, jump_type, size))

    def __exit__(self, exc_type, exc_value, traceback):
        if exc_type is None:
            (start, jump_type, size) = self._jumps.pop()
            if jump_type:
                self.read(size - (self._pos - start))
            else:
                align = (self._pos - start) % size
                if align:
                    self.read(size - align)

    def unpacks(self, format, size):
        return struct.unpack(format, self.read(size))

    def unpack(self, format, size):
        return self.unpacks(format, size)[0]

    def raise_error(self, reason, back=0):
        raise OleFormParsingError('{0}:{1}: {2}'.format(self._path, self._pos - back, reason))

    def check_values(self, name, format, size, expected):
        value = self.unpacks(format, size)
        if value != expected:
            self.raise_error('Invalid {0}: expected {1} got {2}'.format(name, str(expected), str(value)))

    def check_value(self, name, format, size, expected):
        self.check_values(name, format, size, (expected,))


def consume_TextProps(stream):
    # TextProps: [MS-OFORMS] 2.3.1
    stream.check_values('TextProps (versions)', '<BB', 2, (0, 2))
    cbTextProps = stream.unpack('<H', 2)
    stream.read(cbTextProps)

def consume_GuidAndFont(stream):
    # GuidAndFont: [MS-OFORMS] 2.4.7
    UUIDS = stream.unpacks('<LHH', 8) + stream.unpacks('>Q', 8)
    if UUIDS == (199447043, 36753, 4558, 11376937813817407569):
        # UUID == {0BE35203-8F91-11CE-9DE300AA004BB851}
        # StdFont: [MS-OFORMS] 2.4.12
        stream.check_value('StdFont (version)', '<B', 1, 1)
        # Skip sCharset, bFlags, sWeight, ulHeight
        stream.read(9)
        bFaceLen = stream.unpack('<B', 1)
        stream.read(bFaceLen)
    elif UUIDS == (2948729120, 55886, 4558, 13349514450607572916):
        # UUID == {AFC20920-DA4E-11CE-B94300AA006887B4}
        consume_TextProps(stream)
    else:
        stream.raise_error('Invalid GuidAndFont (UUID)', 16)

def consume_GuidAndPicture(stream):
    # GuidAndPicture: [MS-OFORMS] 2.4.8
    # UUID == {0BE35204-8F91-11CE-9DE3-00AA004BB851}
    stream.check_values('GuidAndPicture (UUID part 1)', '<LHH', 8, (199447044, 36753, 4558))
    stream.check_value('GuidAndPicture (UUID part 1)', '>Q', 8, 11376937813817407569)
    # StdPicture: [MS-OFORMS] 2.4.13
    stream.check_value('StdPicture (Preamble)', '<L', 4, 0x0000746C)
    size = stream.unpack('<L', 4)
    stream.read(size)

def consume_CountOfBytesWithCompressionFlag(stream):
    # CountOfBytesWithCompressionFlag or CountOfCharsWithCompressionFlag: [MS-OFORMS] 2.4.14.2 or 2.4.14.3
    count = stream.unpack('<L', 4)
    return count & 0x7FFFFFFF

def consume_SiteClassInfo(stream):
   # SiteClassInfo: [MS-OFORMS] 2.2.10.10.1
   stream.check_value('SiteClassInfo (version)', '<H', 2, 0)
   cbClassTable = stream.unpack('<H', 2)
   stream.read(cbClassTable)

def consume_FormObjectDepthTypeCount(stream):
   # FormObjectDepthTypeCount: [MS-OFORMS] 2.2.10.7
   (depth, mixed) = stream.unpacks('<BB', 2)
   if mixed & 0x80:
       stream.check_value('FormObjectDepthTypeCount (SITE_TYPE)', '<B', 1, 1)
       return mixed ^ 0x80
   if mixed != 1:
       stream.raise_error('Invalid FormObjectDepthTypeCount (SITE_TYPE): expected 1 got {0}'.format(str(mixed)))
   return 1

def consume_OleSiteConcreteControl(stream):
    # OleSiteConcreteControl: [MS-OFORMS] 2.2.10.12.1
    stream.check_value('OleSiteConcreteControl (version)', '<H', 2, 0)
    cbSite = stream.unpack('<H', 2)
    with stream.will_jump_to(cbSite):
        propmask = SitePropMask(stream.unpack('<L', 4))
        # SiteDataBlock: [MS-OFORMS] 2.2.10.12.3
        name_len = tag_len = id = 0
        if propmask.fName:
            name_len = consume_CountOfBytesWithCompressionFlag(stream)
        if propmask.fTag:
            tag_len = consume_CountOfBytesWithCompressionFlag(stream)
        if propmask.fID:
            id = stream.unpack('<L', 4)
        for prop in ['fHelpContextID', 'fBitFlags', 'fObjectStreamSize']:
            if propmask[prop]:
                stream.read(4)
        tabindex = ClsidCacheIndex = 0
        with stream.will_pad():
            if propmask.fTabIndex:
                tabindex = stream.unpack('<H', 2)
            if propmask.fClsidCacheIndex:
                ClsidCacheIndex = stream.unpack('<H', 2)
            if propmask.fGroupID:
                stream.read(2)
        # For the next 4 entries, the documentation adds padding, but it should already be aligned??
        for prop in ['fControlTipText', 'fRuntimeLicKey', 'fControlSource', 'fRowSource']:
            if propmask[prop]:
                stream.read(4)
        # SiteExtraDataBlock: [MS-OFORMS] 2.2.10.12.4
        name = stream.read(name_len)
        tag = stream.read(tag_len)
        return {'name': name, 'tag': tag, 'id': id, 'tabindex': tabindex,
               'ClsidCacheIndex': ClsidCacheIndex, 'value': None}

def consume_FormControl(stream):
    # FormControl: [MS-OFORMS] 2.2.10.1
    stream.check_values('FormControl (versions)', '<BB', 2, (0, 4))
    cbform = stream.unpack('<H', 2)
    with stream.will_jump_to(cbform):
        propmask = FormPropMask(stream.unpack('<L', 4))
        # FormDataBlock: [MS-OFORMS] 2.2.10.3
        for prop in ['fBackColor', 'fForeColor', 'fNextAvailableID']:
            if propmask[prop]:
                stream.read(4)
        if propmask.fBooleanProperties:
            BooleanProperties = stream.unpack('<L', 4)
            FORM_FLAG_DONTSAVECLASSTABLE = (BooleanProperties & (1<<15)) >> 15
        else:
            FORM_FLAG_DONTSAVECLASSTABLE = 0
        # Skip the rest of DataBlock and ExtraDataBlock
    # FormStreamData: [MS-OFORMS] 2.2.10.5
    if propmask.fMouseIcon:
        consume_GuidAndPicture(stream)
    if propmask.fFont:
        consume_GuidAndFont(stream)
    if propmask.fPicture:
        consume_GuidAndPicture(stream)
    # FormSiteData: [MS-OFORMS] 2.2.10.6
    if not FORM_FLAG_DONTSAVECLASSTABLE:
        CountOfSiteClassInfo = stream.unpack('<H', 2)
        for i in range(CountOfSiteClassInfo):
            consume_SiteClassInfo(stream)
    (CountOfSites, CountOfBytes) = stream.unpacks('<LL', 8)
    remaining_SiteDepthsAndTypes = CountOfSites
    with stream.will_pad():
        while remaining_SiteDepthsAndTypes > 0:
            remaining_SiteDepthsAndTypes -= consume_FormObjectDepthTypeCount(stream)
    for i in range(CountOfSites):
        yield consume_OleSiteConcreteControl(stream)

def consume_MorphDataControl(stream):
    # MorphDataControl: [MS-OFORMS] 2.2.5.1
    stream.check_values('MorphDataControl (versions)', '<BB', 2, (0, 2))
    cbMorphData = stream.unpack('<H', 2)
    with stream.will_jump_to(cbMorphData):
        propmask = MorphDataPropMask(stream.unpack('<Q', 8))
        # MorphDataDataBlock: [MS-OFORMS] 2.2.5.3
        for prop in ['fVariousPropertyBits', 'fBackColor', 'fForeColor', 'fMaxLength']:
            if propmask[prop]:
                stream.read(4)
        with stream.will_pad():
            for prop in ['fBorderStyle', 'fScrollBars', 'fDisplayStyle', 'fMousePointer']:
                if propmask[prop]:
                    stream.read(1)
        # PasswordChar, BoundColumn, TextColumn, ColumnCount, and ListRows are 2B + pad = 4B
        # ListWidth is 4B + pad = 4B
        for prop in ['fPasswordChar', 'fListWidth', 'fBoundColumn', 'fTextColumn', 'fColumnCount',
                     'fListRows']:
            if propmask[prop]:
                stream.read(4)
        with stream.will_pad():
            if propmask.fcColumnInfo:
                stream.read(2)
            for prop in ['fMatchEntry', 'fListStyle', 'fShowDropButtonWhen', 'fDropButtonStyle',
                         'fMultiSelect']:
                if propmask[prop]:
                    stream.read(1)
        if propmask.fValue:
            value_size = consume_CountOfBytesWithCompressionFlag(stream)
        else:
            value_size = 0
        # Caption, PicturePosition, BorderColor, SpecialEffect, GroupName  are 4B + pad = 4B
        # MouseIcon, Picture, Accelerator are 2B + pad = 4B
        for prop in ['fCaption', 'fPicturePosition', 'fBorderColor', 'fSpecialEffect',
                     'fMouseIcon', 'fPicture', 'fAccelerator', 'fGroupName']:
            if propmask[prop]:
                stream.read(4)
        # MorphDataExtraDataBlock: [MS-OFORMS] 2.2.5.4
        stream.read(8)
        value = stream.read(value_size)
    # MorphDataStreamData: [MS-OFORMS] 2.2.5.5
    if propmask.fMouseIcon:
        consume_GuidAndPicture(stream)
    if propmask.fPicture:
        consume_GuidAndPicture(stream)
    consume_TextProps(stream)
    return value

def consume_ImageControl(stream):
    # ImageControl: [MS-OFORMS] 2.2.3.1
    stream.check_values('ImageControl (versions)', '<BB', 2, (0, 2))
    cbImage = stream.unpack('<H', 2)
    with stream.will_jump_to(cbImage):
        propmask = ImagePropMask(stream.unpack('<L', 4))
        # Skip the DataBlock and the ExtraDataBlock
    # ImageStreamData: [MS-OFORMS] 2.2.3.5
    if propmask.fPicture:
        consume_GuidAndPicture(stream)
    if propmask.fMouseIcon:
        consume_GuidAndPicture(stream)

def consume_CommandButtonControl(stream):
    # CommandButtonControl: [MS-OFORMS] 2.2.1.1
    stream.check_values('CommandButtonControl (versions)', '<BB', 2, (0, 2))
    cbCommandButton = stream.unpack('<H', 2)
    with stream.will_jump_to(cbCommandButton):
        propmask = CommandButtonPropMask(stream.unpack('<L', 4))
        # Skip the DataBlock and the ExtraDataBlock
    # ImageStreamData: [MS-OFORMS] 2.2.1.5
    if propmask.fPicture:
        consume_GuidAndPicture(stream)
    if propmask.fMouseIcon:
        consume_GuidAndPicture(stream)
    consume_TextProps(stream)

def consume_SpinButtonControl(stream):
    # SpinButtonControl: [MS-OFORMS] 2.2.8.1
    stream.check_values('SpinButtonControl (versions)', '<BB', 2, (0, 2))
    cbSpinButton = stream.unpack('<H', 2)
    with stream.will_jump_to(cbSpinButton):
         propmask = SpinButtonPropMask(stream.unpack('<L', 4))
        # Skip the DataBlock and the ExtraDataBlock
    # SpinButtonStreamData: [MS-OFORMS] 2.2.8.5
    if propmask.fMouseIcon:
        consume_GuidAndPicture(stream)

def consume_TabStripControl(stream):
    # TabStripControl: [MS-OFORMS] 2.2.9.1
    stream.check_values('TabStripControl (versions)', '<BB', 2, (0, 2))
    cbTabStrip = stream.unpack('<H', 2)
    with stream.will_jump_to(cbTabStrip):
        propmask = TabStripPropMask(stream.unpack('<L', 4))
        # TabStripDataBlock: [MS-OFORMS] 2.2.9.3
        # All are 4B (or 4B + pad = 4B), except MousePointer, which is 1B + pad = 4b
        for prop in ['fListIndex', 'fBackColor', 'fForeColor', 'fSize', 'fTabOrientation'
                     'fTabStyle', 'fTabFixedWidth', 'fTabFixedHeight', 'fTipStrings',
                     'fNames', 'fVariousPropertyBits', 'fTabsAllocated', 'fTags']:
            if propmask[prop]:
                stream.read(4)
        tabData = 0
        if propmask['fTabData']:
            tabData = stream.unpack('<L', 4)
        # Skip the ExtraDataBlock
    # TabStripStreamData: [MS-OFORMS] 2.2.9.5
    if propmask.fMouseIcon:
        consume_GuidAndPicture(stream)
    # TextProps
    consume_TextProps(stream)
    # TabStripTabFlagData: [MS-OFORMS] 2.2.9.6
    for i in range(tabData):
         stream.read(4)

def consume_LabelControl(stream):
    # LabelControl: [MS-OFORMS] 2.2.4.1
    stream.check_values('LabelControl (versions)', '<BB', 2, (0, 2))
    cbLabel = stream.unpack('<H', 2)
    with stream.will_jump_to(cbLabel):
        propmask = LabelPropMask(stream.unpack('<L', 4))
        # Skip the DataBlock and the ExtraDataBlock
    # LabelStreamData: [MS-OFORMS] 2.2.4.5
    if propmask.fPicture:
        consume_GuidAndPicture(stream)
    if propmask.fMouseIcon:
        consume_GuidAndPicture(stream)
    # TextProps
    consume_TextProps(stream)

def consume_ScrollBarControl(stream):
    # ScrollBarControl: [MS-OFORMS] 2.2.7.1
    stream.check_values('LabelControl (versions)', '<BB', 2, (0, 2))
    cbScrollBar = stream.unpack('<H', 2)
    with stream.will_jump_to(cbScrollBar):
        propmask = ScrollBarPropMask(stream.unpack('<L', 4))
        # Skip the DataBlock and the ExtraDataBlock
    # ScrollBarStreamData: [MS-OFORMS] 2.2.7.5
    if propmask.fMouseIcon:
        consume_GuidAndPicture(stream)

def extract_OleFormVariables(ole_file, stream_dir):
    control = ExtendedStream.open(ole_file, '/'.join(stream_dir + ['f']))
    variables = list(consume_FormControl(control))
    # print('/'.join(stream_dir + ['o']))
    data = ExtendedStream.open(ole_file, '/'.join(stream_dir + ['o']))
    for var in variables:
        # See FormEmbeddedActiveXControlCached for type definition: [MS-OFORMS] 2.4.5
        if var['ClsidCacheIndex'] == 7:
            raise OleFormParsingError('Malformed document: Forms should be stored in the f stream')
        elif var['ClsidCacheIndex'] == 12:
            consume_ImageControl(data)
        elif var['ClsidCacheIndex'] == 14:
            raise OleFormParsingError('Malformed document: Frames should be stored in the f stream')
        elif var['ClsidCacheIndex'] in [15, 23, 24, 25, 26, 27, 28]:
            var['value'] = consume_MorphDataControl(data)
        elif var['ClsidCacheIndex'] == 16:
            consume_SpinButtonControl(data)
        elif var['ClsidCacheIndex'] == 17:
            consume_CommandButtonControl(data)
        elif var['ClsidCacheIndex'] == 18:
            consume_TabStripControl(data)
        elif var['ClsidCacheIndex'] == 21:
            consume_LabelControl(data)
        elif var['ClsidCacheIndex'] == 47:
            consume_ScrollBarControl(data)
        elif var['ClsidCacheIndex'] == 57:
            raise OleFormParsingError('Malformed document: MultiPages should be stored in a x stream')
        else:
            # TODO: use logging instead of print
            print('ERROR: Unsupported stored type in user form: {0}'.format(str(var['ClsidCacheIndex'])))
            break
    return variables
