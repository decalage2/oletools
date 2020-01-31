#!/usr/bin/env python
"""
oleform.py

oleform is a python module to parse VBA forms in Microsoft Office files.

Authors: see https://github.com/decalage2/oletools/commits/master/oletools/oleform.py
License: BSD, see source code or documentation

oleform is part of the python-oletools package:
http://www.decalage.info/python/oletools
"""

# === LICENSE ==================================================================

# oletools is copyright (c) 2012-2020 Philippe Lagadec (http://www.decalage.info)
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


# REFERENCES:
# - MS-OFORMS: https://msdn.microsoft.com/en-us/library/office/cc313125%28v=office.12%29.aspx?f=255&MSPPError=-2147217396

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

    def consume(self, stream, props):
        for (name, size) in props:
            if self[name]:
                stream.read(size)

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
        self._padding = False
        self._pad_start = 0

    @classmethod
    def open(cls, ole_file, path):
        stream = ole_file.openstream(path)
        return cls(stream, path)

    def _read(self, size):
        self._pos += size
        return self._stream.read(size)

    def _pad(self, start, size=4):
        offset = (self._pos - start) % size
        if offset:
            self._read(size - offset)

    def read(self, size):
        if self._padding:
            self._pad(self._pad_start, size)
        return self._read(size)

    def will_jump_to(self, size):
        self._next_jump = ('jump', (self._pos, size))
        return self

    def will_pad(self):
        self._next_jump = ('pad', self._pos)
        return self

    def padded_struct(self):
        self._next_jump = ('padded', (self._padding, self._pad_start))
        self._padding = True
        self._pad_start = self._pos
        return self

    def __enter__(self):
        assert(self._next_jump)
        self._jumps.append(self._next_jump)
        self._next_jump = None

    def __exit__(self, exc_type, exc_value, traceback):
        if exc_type is None:
            (jump_type, data) = self._jumps.pop()
            if jump_type == 'jump':
                (start, size) = data
                consummed = self._pos - start
                if consummed > size:
                    self.raise_error('Bad jump: too much read ({0} > {1})'.format(consummed, size))
                self.read(size - consummed)
            elif jump_type == 'pad':
                self._pad(data)
            elif jump_type == 'padded':
                (prev_padding, prev_pad_start) = data
                self._pad(self._pad_start)
                self._padding = prev_padding
                self._pad_start = prev_pad_start

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
        with stream.padded_struct():
            name_len = tag_len = id = 0
            if propmask.fName:
                name_len = consume_CountOfBytesWithCompressionFlag(stream)
            if propmask.fTag:
                tag_len = consume_CountOfBytesWithCompressionFlag(stream)
            if propmask.fID:
                id = stream.unpack('<L', 4)
            propmask.consume(stream, [('fHelpContextID', 4), ('fBitFlags', 4), ('fObjectStreamSize', 4)])
            tabindex = ClsidCacheIndex = 0
            if propmask.fTabIndex:
                tabindex = stream.unpack('<H', 2)
            if propmask.fClsidCacheIndex:
                ClsidCacheIndex = stream.unpack('<H', 2)
            if propmask.fGroupID:
                stream.read(2)
            # Get the size of the ControlTipText, if needed.
            control_tip_text_len = 0
            if propmask.fControlTipText:
                control_tip_text_len = consume_CountOfBytesWithCompressionFlag(stream)
            propmask.consume(stream, [('fRuntimeLicKey', 4), ('fControlSource', 4), ('fRowSource', 4)])
        # SiteExtraDataBlock: [MS-OFORMS] 2.2.10.12.4
        name = None
        if (name_len > 0):
            name = stream.read(name_len)
        # Consume 2 null bytes between name and tag.
        #if ((tag_len > 0) or (control_tip_text_len > 0)):
        #    stream.read(2)
        #    # Sometimes it looks like 2 extra null bytes go here whether or not there is a tag.
        tag = None
        if (tag_len > 0):
            tag = stream.read(tag_len)
        # Skip SitePosition.
        if propmask.fPosition:
            stream.read(8)
        control_tip_text = stream.read(control_tip_text_len)
        if (len(control_tip_text) == 0):
            control_tip_text = None
        return {'name': name, 'tag': tag, 'id': id, 'tabindex': tabindex,
                'ClsidCacheIndex': ClsidCacheIndex, 'value': None, 'caption': None,
                'control_tip_text':control_tip_text}

def consume_FormControl(stream):
    # FormControl: [MS-OFORMS] 2.2.10.1
    stream.check_values('FormControl (versions)', '<BB', 2, (0, 4))
    cbform = stream.unpack('<H', 2)
    with stream.will_jump_to(cbform):
        propmask = FormPropMask(stream.unpack('<L', 4))
        # FormDataBlock: [MS-OFORMS] 2.2.10.3
        propmask.consume(stream, [('fBackColor', 4), ('fForeColor', 4), ('fNextAvailableID', 4)])
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
    with stream.will_jump_to(CountOfBytes):
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
        with stream.padded_struct():
            propmask.consume(stream, [('fVariousPropertyBits', 4), ('fBackColor', 4),
                                      ('fForeColor', 4), ('fMaxLength', 4),
                                      ('fBorderStyle', 1), ('fScrollBars', 1),
                                      ('fDisplayStyle', 1), ('fMousePointer', 1),
                                      ('fPasswordChar', 2), ('fListWidth', 4),
                                      ('fBoundColumn', 2), ('fTextColumn', 2),
                                      ('fColumnCount', 2), ('fListRows', 2),
                                      ('fcColumnInfo', 2), ('fMatchEntry', 1),
                                      ('fListStyle', 1), ('fShowDropButtonWhen', 1),
                                      ('fDropButtonStyle', 1), ('fMultiSelect', 1)])
            if propmask.fValue:
                value_size = consume_CountOfBytesWithCompressionFlag(stream)
            else:
                value_size = 0
            if propmask.fCaption:
                caption_size = consume_CountOfBytesWithCompressionFlag(stream)
            else:
                caption_size = 0
            propmask.consume(stream, [('fPicturePosition', 4),
                                      ('fBorderColor', 4), ('fSpecialEffect', 4),
                                      ('fMouseIcon', 2), ('fPicture', 2),
                                      ('fAccelerator', 2)])
            if propmask.fGroupName:
                group_name_size = consume_CountOfBytesWithCompressionFlag(stream)
            else:
                group_name_size = 0
        # MorphDataExtraDataBlock: [MS-OFORMS] 2.2.5.4
        # Discard Size
        stream.read(8)
        value = stream.read(value_size)
        # Read caption text.
        caption = ""
        if (caption_size > 0):
            caption = stream.read(caption_size)
        # Read groupname text.
        group_name = ""
        if (group_name_size > 0):
            group_name = stream.read(group_name_size)
            
    # MorphDataStreamData: [MS-OFORMS] 2.2.5.5
    if propmask.fMouseIcon:
        consume_GuidAndPicture(stream)
    if propmask.fPicture:
        consume_GuidAndPicture(stream)
    consume_TextProps(stream)
    return (value, caption, group_name)

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
        propmask.consume(stream, [('fListIndex', 4), ('fBackColor', 4),
                                  ('fForeColor', 4), ('fSize', 4),
                                  ('fMousePointer', 1), ('fTabOrientation', 4),
                                  ('fTabStyle', 4), ('fTabFixedWidth', 4),
                                  ('fTabFixedHeight', 4), ('fTipStrings', 4),
                                  ('fNames', 4), ('fVariousPropertyBits', 4),
                                  ('fTabsAllocated', 4), ('fTags', 4)])
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
        # LabelDataBlock: [MS-OFORMS] 2.2.4.3
        with stream.padded_struct():
            propmask.consume(stream, [('fForeColor', 4), ('fBackColor', 4),
                                      ('fVariousPropertyBits', 4)])
            if propmask.fCaption:
                caption_size = consume_CountOfBytesWithCompressionFlag(stream)
            else:
                caption_size = 0
            propmask.consume(stream, [('fPicturePosition', 4), ('fMousePointer', 1),
                                      ('fBorderColor', 4), ('fBorderStyle', 2),
                                      ('fSpecialEffect', 2), ('fPicture', 2),
                                      ('fAccelerator', 2), ('fMouseIcon', 2)])
        # LabelExtraDataBlock: [MS-OFORMS] 2.2.4.4
        caption = stream.read(caption_size)
        stream.read(8)
    # LabelStreamData: [MS-OFORMS] 2.2.4.5
    if propmask.fPicture:
        consume_GuidAndPicture(stream)
    if propmask.fMouseIcon:
        consume_GuidAndPicture(stream)
    # TextProps
    consume_TextProps(stream)
    return caption

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
    data = ExtendedStream.open(ole_file, '/'.join(stream_dir + ['o']))
    for var in variables:
        # See FormEmbeddedActiveXControlCached for type definition: [MS-OFORMS] 2.4.5
        if var['ClsidCacheIndex'] == 7:
            consume_FormControl(data)
        elif var['ClsidCacheIndex'] == 12:
            consume_ImageControl(data)
        elif var['ClsidCacheIndex'] == 14:
            consume_FormControl(data)
        elif var['ClsidCacheIndex'] in [15, 23, 24, 25, 26, 27, 28]:
            var['value'], var['caption'], var['group_name'] = consume_MorphDataControl(data)
        elif var['ClsidCacheIndex'] == 16:
            consume_SpinButtonControl(data)
        elif var['ClsidCacheIndex'] == 17:
            consume_CommandButtonControl(data)
        elif var['ClsidCacheIndex'] == 18:
            consume_TabStripControl(data)
        elif var['ClsidCacheIndex'] == 21:
            var['caption'] = consume_LabelControl(data)
        elif var['ClsidCacheIndex'] == 47:
            consume_ScrollBarControl(data)
        elif var['ClsidCacheIndex'] == 57:
            consume_FormControl(data)
        else:
            # TODO: use logging instead of print
            print('ERROR: Unsupported stored type in user form: {0}'.format(str(var['ClsidCacheIndex'])))
            break
    return variables
