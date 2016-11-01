# xxxswf.py was created by alexander dot hanel at gmail dot com
# version 0.1 
# Date - 12-07-2011 
# To do list
#   - Tag Parser
#   - ActionScript Decompiler

# 2016-11-01 PL: - A few changes for Python 2+3 compatibility

import fnmatch 
import hashlib
import imp
import math
import os
import re
import struct
import sys
import time
from io import BytesIO
from optparse import OptionParser
import zlib

def checkMD5(md5):
# checks if MD5 has been seen in MD5 Dictionary 
# MD5Dict contains the MD5 and the CVE
# For { 'MD5':'CVE', 'MD5-1':'CVE-1', 'MD5-2':'CVE-2'}
    MD5Dict = {'c46299a5015c6d31ad5766cb49e4ab4b':'CVE-XXXX-XXXX'}
    if MD5Dict.get(md5):
        print('\t[BAD] MD5 Match on', MD5Dict.get(md5))
    return    

def bad(f):
    for idx, x in enumerate(findSWF(f)):
        tmp = verifySWF(f,x)
        if tmp != None:
            yaraScan(tmp)
            checkMD5(hashBuff(tmp))
    return 
    
def yaraScan(d):
# d = buffer of the read file 
# Scans SWF using Yara
    # test if yara module is installed
    # if not Yara can be downloaded from http://code.google.com/p/yara-project/
    try:
        imp.find_module('yara')
        import yara 
    except ImportError:
        print('\t[ERROR] Yara module not installed - aborting scan')
        return
    # test for yara compile errors
    try:
        r = yara.compile(r'rules.yar')
    except:
        pass
        print('\t[ERROR] Yara compile error - aborting scan')
        return
    # get matches
    m = r.match(data=d)
    # print matches
    for X in m:
        print('\t[BAD] Yara Signature Hit: %s' % X)
    return

def findSWF(d):
# d = buffer of the read file 
# Search for SWF Header Sigs in files
    return [tmp.start() for tmp in re.finditer(b'CWS|FWS', d.read())]

def hashBuff(d):
# d = buffer of the read file 
# This function hashes the buffer
# source: http://stackoverflow.com/q/5853830
    if type(d) is str:
      d = BytesIO(d)
    md5 = hashlib.md5()
    while True:
        data = d.read(128)
        if not data:
            break
        md5.update(data)
    return md5.hexdigest()

def verifySWF(f,addr):
    # Start of SWF
    f.seek(addr)
    # Read Header
    header = f.read(3)
    # Read Version
    ver = struct.unpack('<b', f.read(1))[0]
    # Read SWF Size
    size = struct.unpack('<i', f.read(4))[0]
    # Start of SWF
    f.seek(addr)
    try:
        # Read SWF into buffer. If compressed read uncompressed size. 
        t = f.read(size)
    except:
        pass
        # Error check for invalid SWF
        print(' - [ERROR] Invalid SWF Size')
        return None
    if type(t) is str:
      f = BytesIO(t)
    # Error check for version above 20
    if ver > 20:
        print(' - [ERROR] Invalid SWF Version')
        return None
    
    if b'CWS' in header:
        try:
            f.read(3)
            tmp = b'FWS' + f.read(5) + zlib.decompress(f.read())
            print(' - CWS Header')
            return tmp
        
        except:
            pass
            print('- [ERROR]: Zlib decompression error. Invalid CWS SWF')
            return None
        
    elif b'FWS' in header:
        try:
            tmp = f.read(size)
            print(' - FWS Header')
            return tmp
        
        except:
            pass
            print(' - [ERROR] Invalid SWF Size')
            return None
        
    else:
        print(' - [Error] Logic Error Blame Programmer')
        return None
    
def headerInfo(f):
# f is the already opended file handle 
# Yes, the format is is a rip off SWFDump. Can you blame me? Their tool is awesome.
    # SWFDump FORMAT    
    # [HEADER]        File version: 8
    # [HEADER]        File is zlib compressed. Ratio: 52%
    # [HEADER]        File size: 37536
    # [HEADER]        Frame rate: 18.000000
    # [HEADER]        Frame count: 323
    # [HEADER]        Movie width: 217.00
    # [HEADER]        Movie height: 85.00
    if type(f) is str:
      f = BytesIO(f)
    sig = f.read(3)             
    print('\t[HEADER] File header: %s' % sig)
    if b'C' in sig:
        print('\t[HEADER] File is zlib compressed.')
    version = struct.unpack('<b', f.read(1))[0]
    print('\t[HEADER] File version: %d' % version)
    size = struct.unpack('<i', f.read(4))[0]
    print('\t[HEADER] File size: %d' % size)
    # deflate compressed SWF
    if b'C' in sig:
        f = verifySWF(f,0)
        if type(f) is str:
            f = BytesIO(f)
        f.seek(0, 0)
        x = f.read(8)
    ta = f.tell()
    tmp = struct.unpack('<b', f.read(1))[0]
    nbit =  tmp >> 3
    print('\t[HEADER] Rect Nbit: %d' % nbit)
    # Curretely the nbit is static at 15. This could be modified in the
    # future. If larger than 9 this will break the struct unpack. Will have
    # to revist must be a more effective way to deal with bits. Tried to keep
    # the algo but damn this is ugly...
    f.seek(ta)
    rect =  struct.unpack('>Q', f.read(int(math.ceil((nbit*4)/8.0))))[0]
    tmp = struct.unpack('<b', f.read(1))[0]
    tmp = bin(tmp>>7)[2:].zfill(1)
    # bin requires Python 2.6 or higher
    # skips string '0b' and the nbit 
    rect =  bin(rect)[7:] 
    xmin = int(rect[0:nbit-1],2)
    print('\t[HEADER] Rect Xmin: %d' % xmin)
    xmax = int(rect[nbit:(nbit*2)-1],2)
    print('\t[HEADER] Rect Xmax: %d' % xmax)
    ymin = int(rect[nbit*2:(nbit*3)-1],2)
    print('\t[HEADER] Rect Ymin: %d' % ymin)
    # one bit needs to be added, my math might be off here
    ymax = int(rect[nbit*3:(nbit*4)-1] + str(tmp) ,2)
    print('\t[HEADER] Rect Ymax: %d' % ymax)
    framerate = struct.unpack('<H', f.read(2))[0]
    print('\t[HEADER] Frame Rate: %d' % framerate)
    framecount = struct.unpack('<H', f.read(2))[0] 
    print('\t[HEADER] Frame Count: %d' % framecount)
       
def walk4SWF(path):
    # returns a list of [folder-path, [addr1,addrw2]]
    # Don't ask, will come back to this code. 
    p = ['',[]]
    r = p*0
    if os.path.isdir(path) != True and path != '':
        print('\t[ERROR] walk4SWF path must be a dir.')
        return 
    for root, dirs, files in os.walk(path):
        for name in files:
            try: 
                x = open(os.path.join(root, name), 'rb')
            except:
                pass
                break
            y = findSWF(x)
            if len(y) != 0:
                # Path of file SWF
                p[0] = os.path.join(root, name)
                # contains list of the file offset of SWF header
                p[1] = y
                r.insert(len(r),p)
                p = ['',[]]
                y = ''
            x.close()
    return r

def tagsInfo(f):
    return

def fileExist(n, ext):
    # Checks the working dir to see if the file is
    # already in the dir. If exists the file will
    # be named name.count.ext (n.c.ext). No more than
    # 50 matching MD5s will be written to the dir. 
    if os.path.exists( n + '.' + ext):
                c = 2
                while os.path.exists(n + '.' + str(c) + '.' + ext):
                    c =  c + 1
                    if c == 50:
                        print('\t[ERROR] Skipped 50 Matching MD5 SWFs')
                        break
                n = n + '.' + str(c)
                
    return n + '.' + ext
    
def CWSize(f):
    # The file size in the header is of the uncompressed SWF.
    # To estimate the size of the compressed data, we can grab
    # the length, read that amount, deflate the data, then
    # compress the data again, and then call len(). This will
    # give us the length of the compressed SWF. 
    return

def compressSWF(f):
    if type(f) is str:
      f = BytesIO(f)
    try:
        f.read(3)
        tmp = b'CWS' + f.read(5) + zlib.compress(f.read())
        return tmp
    except:
        pass
        print('\t[ERROR] SWF Zlib Compression Failed')
        return None

def disneyland(f,filename, options):
    # because this is where the magic happens
    # but seriously I did the recursion part last..
    retfindSWF = findSWF(f)
    f.seek(0)
    print('\n[SUMMARY] %d SWF(s) in MD5:%s:%s' % ( len(retfindSWF),hashBuff(f), filename ))
    # for each SWF in file 
    for idx, x in enumerate(retfindSWF):
        print('\t[ADDR] SWF %d at %s' % (idx+1, hex(x)))
        f.seek(x)
        h = f.read(1)
        f.seek(x)
        swf = verifySWF(f,x)
        if swf == None:
            continue
        if options.extract != None:
            name = fileExist(hashBuff(swf), 'swf')
            print('\t\t[FILE] Carved SWF MD5: %s' % name)
            try:
                o = open(name, 'wb+')
            except IOError as e:
                print('\t[ERROR] Could Not Create %s ' % e)
                continue 
            o.write(swf)
            o.close()
        if options.yara != None:
            yaraScan(swf)
        if options.md5scan != None:
            checkMD5(hashBuff(swf))
        if options.decompress != None:
            name = fileExist(hashBuff(swf), 'swf')
            print('\t\t[FILE] Carved SWF MD5: %s' % name)
            try:
                o = open(name, 'wb+')
            except IOError as e:
                print('\t[ERROR] Could Not Create %s ' % e)
                continue
            o.write(swf)
            o.close()
        if options.header != None:
            headerInfo(swf)
        if options.compress != None:
            swf = compressSWF(swf)
            if swf == None:
                continue 
            name = fileExist(hashBuff(swf), 'swf')
            print('\t\t[FILE] Compressed SWF MD5: %s' % name)
            try:
                o = open(name, 'wb+')
            except IOError as e:
                print('\t[ERROR] Could Not Create %s ' % e)
                continue
            o.write(swf)
            o.close()

def main():
    # Scenarios:
    # Scan file for SWF(s)
    # Scan file for SWF(s) and extract them 
    # Scan file for SWF(s) and scan them with Yara
    # Scan file for SWF(s), extract them and scan with Yara
    # Scan directory recursively for files that contain SWF(s) 
    # Scan directory recursively for files that contain SWF(s) and extract them
    
    parser = OptionParser()
    usage = 'usage: %prog [options] <file.bad>'
    parser = OptionParser(usage=usage)
    parser.add_option('-x', '--extract', action='store_true', dest='extract', help='Extracts the embedded SWF(s), names it MD5HASH.swf & saves it in the working dir. No addition args needed')
    parser.add_option('-y', '--yara', action='store_true', dest='yara', help='Scans the SWF(s) with yara. If the SWF(s) is compressed it will be deflated. No addition args needed')
    parser.add_option('-s', '--md5scan', action='store_true', dest='md5scan', help='Scans the SWF(s) for MD5 signatures. Please see func checkMD5 to define hashes. No addition args needed')
    parser.add_option('-H', '--header', action='store_true', dest='header', help='Displays the SWFs file header. No addition args needed')
    parser.add_option('-d', '--decompress', action='store_true', dest='decompress', help='Deflates compressed SWFS(s)')
    parser.add_option('-r', '--recdir', dest='PATH', type='string', help='Will recursively scan a directory for files that contain SWFs. Must provide path in quotes')
    parser.add_option('-c', '--compress', action='store_true', dest='compress', help='Compresses the SWF using Zlib')
    
    (options, args) = parser.parse_args()

    # Print help if no argurments are passed
    if len(sys.argv) < 2:
        parser.print_help()
        return

    # Note files can't start with '-'
    if '-' in sys.argv[len(sys.argv)-1][0] and options.PATH == None:
        parser.print_help()
        return
    
    # Recusive Search
    if options.PATH != None:
        paths = walk4SWF(options.PATH)
        for y in paths:
            #if sys.argv[0] not in y[0]:
            try:
                t = open(y[0], 'rb+')
                disneyland(t, y[0],options)
            except IOError:
                pass
        return 
        
    # try to open file 
    try:
        f = open(sys.argv[len(sys.argv)-1],'rb+')
        filename = sys.argv[len(sys.argv)-1]
    except Exception:
        print('[ERROR] File can not be opended/accessed')
        return

    disneyland(f,filename,options)
    f.close()
    return 
        
if __name__ == '__main__':
   main()
   
