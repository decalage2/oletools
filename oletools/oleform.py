#!/usr/bin/env python

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

class OleUserFormParser(object):
    def __init__(self, control_stream, data_stream):
        self.variables = []
        self.set_stream(control_stream)
        self.consume_FormControl()
        self.set_stream(data_stream)
        self.consume_stored_data()

    def set_stream(self, stream):
        self._pos = 0
        self._frozen_pos = []
        self._stream = stream

    def read(self, size):
        self._pos += size
        return self._stream.read(size)

    def freeze(self):
        self._frozen_pos.append(self._pos)

    def unfreeze(self, size):
        self.read(self._frozen_pos.pop() - self._pos + size)

    def unfreeze_pad(self):
        align_pos = (self._pos - self._frozen_pos.pop()) % 4
        if align_pos:
            self.read(4 - align_pos)

    def unpacks(self, format, size):
        return struct.unpack(format, self.read(size))

    def unpack(self, format, size):
        return self.unpacks(format, size)[0]

    def check_values(self, name, format, size, expected):
        value = self.unpacks(format, size)
        if value != expected:
            raise OleFormParsingError('Invalid {0} at {1}: expected {2} got {3}'.format(name, self._pos - size, str(expected), str(value)))

    def check_value(self, name, format, size, expected):
        self.check_values(name, format, size, (expected,))

    def consume_TextProps(self):
        # TextProps: [MS-OFORMS] 2.3.1
        self.check_values('TextProps (versions)', '<BB', 2, (0, 2))
        cbTextProps = self.unpack('<H', 2)
        self.read(cbTextProps)

    def consume_GuidAndFont(self):
        # GuidAndFont: [MS-OFORMS] 2.4.7
        UUIDS = self.unpacks('<LHH', 8) + self.unpacks('>Q', 8)
        if UUIDS == (199447043, 36753, 4558, 11376937813817407569L):
            # UUID == {0BE35203-8F91-11CE-9DE300AA004BB851}
            # StdFont: [MS-OFORMS] 2.4.12
            self.check_value('StdFont (version)', '<B', 1, 1)
            # Skip sCharset, bFlags, sWeight, ulHeight
            self.read(9)
            bFaceLen = self.unpack('<B', 1)
            self.read(bFaceLen)
        elif UUIDs == (2948729120, 55886, 4558, 13349514450607572916L):
            # UUID == {AFC20920-DA4E-11CE-B94300AA006887B4}
            self.consume_TextProps()
        else:
            raise OleFormParsingError('Invalid GuidAndFont at {0}: UUID'.format(self._pos - 16))

    def consume_GuidAndPicture(self):
        # GuidAndPicture: [MS-OFORMS] 2.4.8
        # UUID == {0BE35204-8F91-11CE-9DE3-00AA004BB851}
        self.check_values('GuidAndPicture (UUID part 1)', '<LHH', 8, (199447044, 36753, 4558))
        self.check_value('GuidAndPicture (UUID part 1)', '>Q', 8, 11376937813817407569L)
        # StdPicture: [MS-OFORMS] 2.4.13
        self.check_value('StdPicture (Preamble)', '<L', 4, 0x0000746C)
        size = self.unpack('<L', 4)
        self.read(size)

    def consume_CountOfBytesWithCompressionFlag(self):
        # CountOfBytesWithCompressionFlag or CountOfCharsWithCompressionFlag: [MS-OFORMS] 2.4.14.2 or 2.4.14.3
        count = self.unpack('<L', 4)
        if not count & 0x80000000 and count != 0:
            raise OleFormParsingError('Uncompress string length at {0}', self._pos - 4)
        return count & 0x7FFFFFFF

    def consume_SiteClassInfo(self):
       # SiteClassInfo: [MS-OFORMS] 2.2.10.10.1
       self.check_value('SiteClassInfo (version)', '<H', 2, 0)
       cbClassTable = self.unpack('<H', 2)
       self.read(cbClassTable)

    def consume_FormObjectDepthTypeCount(self):
       # FormObjectDepthTypeCount: [MS-OFORMS] 2.2.10.7
       (depth, mixed) = self.unpacks('<BB', 2)
       if mixed & 0x80:
           self.check_value('FormObjectDepthTypeCount (SITE_TYPE)', '<B', 1, 1)
           return mixed ^ 0x80
       if mixed != 1:
           raise OleFormParsingError('Invalid FormObjectDepthTypeCount (SITE_TYPE) at {0}: expected 1 got {3}'.format(self._pos - 2, str(mixed)))
       return 1

    def consume_OleSiteConcreteControl(self):
        # OleSiteConcreteControl: [MS-OFORMS] 2.2.10.12.1
        self.check_value('OleSiteConcreteControl (version)', '<H', 2, 0)
        cbSite = self.unpack('<H', 2)
        self.freeze()
        propmask = SitePropMask(self.unpack('<L', 4))
        # SiteDataBlock: [MS-OFORMS] 2.2.10.12.3
        name_len = tag_len = id = 0
        if propmask.fName:
            name_len = self.consume_CountOfBytesWithCompressionFlag()
        if propmask.fTag:
            tag_len = self.consume_CountOfBytesWithCompressionFlag()
        if propmask.fID:
            id = self.unpack('<L', 4)
        for prop in ['fHelpContextID', 'fBitFlags', 'fObjectStreamSize']:
            if propmask[prop]:
                self.read(4)
        tabindex = ClsidCacheIndex = 0
        self.freeze()
        if propmask.fTabIndex:
            tabindex = self.unpack('<H', 2)
        if propmask.fClsidCacheIndex:
            ClsidCacheIndex = self.unpack('<H', 2)
        if propmask.fGroupID:
            self.read(2)
        self.unfreeze_pad()
        # For the next 4 entries, the documentation adds padding, but it should already be aligned??
        for prop in ['fControlTipText', 'fRuntimeLicKey', 'fControlSource', 'fRowSource']:
            if propmask[prop]:
                self.read(4)
        # SiteExtraDataBlock: [MS-OFORMS] 2.2.10.12.4
        name = self.read(name_len)
        tag = self.read(tag_len)
        self.variables.append({'name': name, 'tag': tag, 'id': id,
                               'tabindex': tabindex,
                               'ClsidCacheIndex': ClsidCacheIndex})
        self.unfreeze(cbSite)

    def consume_FormControl(self):
        # FormControl: [MS-OFORMS] 2.2.10.1
        self.check_values('FormControl (versions)', '<BB', 2, (0, 4))
        cbform = self.unpack('<H', 2)
        self.freeze()
        propmask = FormPropMask(self.unpack('<L', 4))
        # FormDataBlock: [MS-OFORMS] 2.2.10.3
        for prop in ['fBackColor', 'fForeColor', 'fNextAvailableID']:
            if propmask[prop]:
                self.read(4)
        if propmask.fBooleanProperties:
            BooleanProperties = self.unpack('<L', 4)
            FORM_FLAG_DONTSAVECLASSTABLE = (BooleanProperties & (1<<15)) >> 15
        else:
            FORM_FLAG_DONTSAVECLASSTABLE = 0
        # Skip the rest of DataBlock and ExtraDataBlock
        self.unfreeze(cbform)
        # FormStreamData: [MS-OFORMS] 2.2.10.5
        if propmask.fMouseIcon:
            self.consume_GuidAndPicture()
        if propmask.fFont:
            self.consume_GuidAndFont()
        if propmask.fPicture:
            self.consume_GuidAndPicture()
        # FormSiteData: [MS-OFORMS] 2.2.10.6
        if not FORM_FLAG_DONTSAVECLASSTABLE:
            CountOfSiteClassInfo = self.unpack('<H', 2)
            for i in range(CountOfSiteClassInfo):
                self.consume_SiteClassInfo()
        (CountOfSites, CountOfBytes) = self.unpacks('<LL', 8)
        remaining_SiteDepthsAndTypes = CountOfSites
        self.freeze()
        while remaining_SiteDepthsAndTypes > 0:
            remaining_SiteDepthsAndTypes -= self.consume_FormObjectDepthTypeCount()
        self.unfreeze_pad()
        for i in range(CountOfSites):
            self.consume_OleSiteConcreteControl()

    def consume_MorphDataControl(self):
        # MorphDataControl: [MS-OFORMS] 2.2.5.1
        self.check_values('MorphDataControl (versions)', '<BB', 2, (0, 2))
        cbMorphData = self.unpack('<H', 2)
        self.freeze()
        propmask = MorphDataPropMask(self.unpack('<Q', 8))
        # MorphDataDataBlock: [MS-OFORMS] 2.2.5.3
        for prop in ['fVariousPropertyBits', 'fBackColor', 'fForeColor', 'fMaxLength']:
            if propmask[prop]:
                self.read(4)
        self.freeze()
        for prop in ['fBorderStyle', 'fScrollBars', 'fDisplayStyle', 'fMousePointer']:
            if propmask[prop]:
                self.read(1)
        self.unfreeze_pad()
        # PasswordChar, BoundColumn, TextColumn, ColumnCount, and ListRows are 2B + pad = 4B
        # ListWidth is 4B + pad = 4B
        for prop in ['fPasswordChar', 'fListWidth', 'fBoundColumn', 'fTextColumn', 'fColumnCount',
                     'fListRows']:
            if propmask[prop]:
                self.read(4)
        self.freeze()
	if propmask.fcColumnInfo:
            self.read(2)
        for prop in ['fMatchEntry', 'fListStyle', 'fShowDropButtonWhen', 'fDropButtonStyle',
                     'fMultiSelect']:
            if propmask[prop]:
                self.read(1)
        self.unfreeze_pad()
        if propmask.fValue:
            value_size = self.consume_CountOfBytesWithCompressionFlag()
        else:
            value_size = 0
        # Caption, PicturePosition, BorderColor, SpecialEffect, GroupName  are 4B + pad = 4B
        # MouseIcon, Picture, Accelerator are 2B + pad = 4B
        for prop in ['fCaption', 'fPicturePosition', 'fBorderColor', 'fSpecialEffect',
                     'fMouseIcon', 'fPicture', 'fAccelerator', 'fGroupName']:
            if propmask[prop]:
                self.read(4)
        # MorphDataExtraDataBlock: [MS-OFORMS] 2.2.5.4
        self.read(8)
        value = self.read(value_size)
        self.unfreeze(cbMorphData)
        # MorphDataStreamData: [MS-OFORMS] 2.2.5.5
        if propmask.fMouseIcon:
            self.consume_GuidAndPicture()
        if propmask.fPicture:
            self.consume_GuidAndPicture()
        self.consume_TextProps()
        return value

    def consume_stored_data(self):
        for var in self.variables:
            if var['ClsidCacheIndex'] != 23:
                raise OleFormParsingError('Unsupported stored type: {0}'.format(str(var['ClsidCacheIndex'])))
            var['value'] = self.consume_MorphDataControl()

def OleFormVariables(ole_file, stream_dir):
    control_stream = ole_file.openstream('/'.join(stream_dir + ['f']))
    data_stream = ole_file.openstream('/'.join(stream_dir + ['o']))
    form = OleUserFormParser(control_stream, data_stream)
    return form.variables
