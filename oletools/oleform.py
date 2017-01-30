#!/usr/bin/env python

import struct

class OleFormParsingError(Exception):
    pass

class Mask(object):
    def __init__(self, val):
        self._val = [(val & (1<<i))>>i for i in range(32)]

    def __str__(self):
        return ', '.join(self._names[i] for i in range(32) if self._val[i])

    def __getattr__(self, name):
        return self._val[self._names.index(name)]

class PropMask(Mask):
    _names = ['Unused1', 'fBackColor', 'fForeColor', 'fNextAvailableID', 'Unused2_0', 'Unused2_1',
              'fBooleanProperties', 'fBooleanProperties', 'fMousePointer', 'fScrollBars',
              'fDisplayedSize', 'fLogicalSize', 'fScrollPosition', 'fGroupCnt', 'Reserved',
              'fMouseIcon', 'fCycle', 'fSpecialEffect', 'fBorderColor', 'fCaption', 'fFont',
              'fPicture', 'fZoom', 'fPictureAlignment', 'fPictureTiling', 'fPictureSizeMode',
              'fShapeCookie', 'fDrawBuffer', 'Unused3_0', 'Unused3_1', 'Unused3_2', 'Unused3_3']

class SitePropMask(Mask):
    _names = ['fName', 'fTag', 'fID', 'fHelpContextID', 'fBitFlags', 'fObjectStreamSize',
              'fTabIndex', 'fClsidCacheIndex', 'fPosition', 'fGroupID', 'Unused1',
              'fControlTipText', 'fRuntimeLicKey', 'fControlSource', 'fRowSource', 'Unused2_0',
              'Unused2_1', 'Unused2_2', 'Unused2_3', 'Unused2_4', 'Unused2_5', 'Unused2_6',
              'Unused2_7', 'Unused2_8', 'Unused2_9', 'Unused2_10', 'Unused2_11', 'Unused2_12',
              'Unused2_13', 'Unused2_14', 'Unused2_15', 'Unused2_16']

class OleUserFormParser(object):
    def __init__(self, stream):
        self.content = []
        self._stream = stream
        self._pos = 0
        self._frozen_pos = []

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
            # TextProps: [MS-OFORMS] 2.3.1
            self.check_value('TextProps (versions)', '<BB', 2, (0, 2))
            cbTextProps = self.unpack('<H', 2)
            self.read(cbTextProps)
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
            print(count)
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
        sitepropmask = SitePropMask(self.unpack('<L', 4))
        # SiteDataBlock: [MS-OFORMS] 2.2.10.12.3
        name_len = tag_len = id = 0
        if sitepropmask.fName:
            name_len = self.consume_CountOfBytesWithCompressionFlag()
        if sitepropmask.fTag:
            tag_len = self.consume_CountOfBytesWithCompressionFlag()
        if sitepropmask.fID:
            id = self.unpack('<L', 4)
        if sitepropmask.fHelpContextID:
            self.read(4)
        if sitepropmask.fBitFlags:
            self.read(4)
        if sitepropmask.fObjectStreamSize:
            self.read(4)
        tabindex = ClsidCacheIndex = 0
        self.freeze()
        if sitepropmask.fTabIndex:
            tabindex = self.unpack('<H', 2)
        if sitepropmask.fClsidCacheIndex:
            ClsidCacheIndex = self.unpack('<H', 2)
        if sitepropmask.fGroupID:
            self.read(2)
        self.unfreeze_pad()
        # For the next 4 entries, the documentation adds padding, but it should already be aligned??
        if sitepropmask.fControlTipText:
            self.read(4)
        if sitepropmask.fRuntimeLicKey:
            self.read(4)
        if sitepropmask.fControlSource:
            self.read(4)
        if sitepropmask.fRowSource:
            self.read(4)
        # SiteExtraDataBlock: [MS-OFORMS] 2.2.10.12.4
        name = self.read(name_len)
        tag = self.read(tag_len)
        self.content.append({'name': name, 'tag': tag, 'id': id,
                          'tabindex': tabindex,
                          'ClsidCacheIndex': ClsidCacheIndex})
        self.unfreeze(cbSite)

    def consume_FormControl(self):
        # FormControl: [MS-OFORMS] 2.2.10.1
        self.check_values('FormControl (versions)', '<BB', 2, (0, 4))
        cbform = self.unpack('<H', 2)
        self.freeze()
        propmask = PropMask(self.unpack('<L', 4))
        # FormDataBlock: [MS-OFORMS] 2.2.10.3
        if propmask.fBackColor:
            self.read(4)
        if propmask.fForeColor:
            self.read(4)
        if propmask.fNextAvailableID:
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

    def consume_stream_o(self):
        # Adapted from plugin_stream_o.py from Didier Stevens's oledump.py
        while(True):
            try:
                (code, length) = self.unpacks('<HH', 4)
            except struct.error:
                break
            self.freeze()
            if code == 0x200:
                fieldtype = self.unpack('<I', 4)
                if fieldtype == 0x80400101:
                    self.read(8)
                    lengthString = self.unpack('<I', 4) & 0x7FFFFFFF #self.consume_CountOfBytesWithCompressionFlag()
                    self.read(8)
                    self.content.append(self.read(lengthString))
                elif fieldtype == 0x80000101:
                    self.content.append('')
            self.unfreeze(length)

def OleFormVariables(ole_file, stream_dir):
    control_stream = ole_file.openstream('/'.join(stream_dir + ['f']))
    control_form = OleUserFormParser(control_stream)
    control_form.consume_FormControl()
    variables = control_form.content
    data_stream = ole_file.openstream('/'.join(stream_dir + ['o']))
    data = OleUserFormParser(data_stream)
    data.consume_stream_o()
    values = data.content
    if len(variables) != len(values):
        raise OleFormParsingError('Incompatible number of variables: {0} VS {1}'.format(len(variables), len(values)))
    for i in range(len(variables)):
        variables[i]['value'] = values[i]
    return variables
