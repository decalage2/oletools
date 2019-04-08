pyxswf
======

pyxswf is a script to detect, extract and analyze Flash objects (SWF files) that may
be embedded in files such as MS Office documents (e.g. Word, Excel),
which is especially useful for malware analysis.

It is part of the [python-oletools](http://www.decalage.info/python/oletools) package.

pyxswf is an extension to [xxxswf.py](http://hooked-on-mnemonics.blogspot.nl/2011/12/xxxswfpy.html) published by Alexander Hanel.

Compared to xxxswf, it can extract streams from MS Office documents by parsing
their OLE structure properly, which is necessary when streams are fragmented.
Stream fragmentation is a known obfuscation technique, as explained on
[http://www.breakingpointsystems.com/resources/blog/evasion-with-ole2-fragmentation/](http://web.archive.org/web/20121118021207/http://www.breakingpointsystems.com/resources/blog/evasion-with-ole2-fragmentation/)

It can also extract Flash objects from RTF documents, by parsing embedded objects encoded in hexadecimal format (-f option).

For this, simply add the -o option to work on OLE streams rather than raw files, or the -f option to work on RTF files.

## Usage

```text
Usage: pyxswf [options] <file.bad>

Options:
  -o, --ole             Parse an OLE file (e.g. Word, Excel) to look for SWF
                        in each stream
  -f, --rtf             Parse an RTF file to look for SWF in each embedded
                        object
  -x, --extract         Extracts the embedded SWF(s), names it MD5HASH.swf &
                        saves it in the working dir. No addition args needed
  -h, --help            show this help message and exit
  -y, --yara            Scans the SWF(s) with yara. If the SWF(s) is
                        compressed it will be deflated. No addition args
                        needed
  -s, --md5scan         Scans the SWF(s) for MD5 signatures. Please see func
                        checkMD5 to define hashes. No addition args needed
  -H, --header          Displays the SWFs file header. No addition args needed
  -d, --decompress      Deflates compressed SWFS(s)
  -r PATH, --recdir=PATH
                        Will recursively scan a directory for files that
                        contain SWFs. Must provide path in quotes
  -c, --compress        Compresses the SWF using Zlib
```

### Example 1 - detecting and extracting a SWF file from a Word document on Windows:

```text
C:\oletools>pyxswf -o word_flash.doc
OLE stream: 'Contents'
[SUMMARY] 1 SWF(s) in MD5:993664cc86f60d52d671b6610813cfd1:Contents
        [ADDR] SWF 1 at 0x8  - FWS Header

C:\oletools>pyxswf -xo word_flash.doc
OLE stream: 'Contents'
[SUMMARY] 1 SWF(s) in MD5:993664cc86f60d52d671b6610813cfd1:Contents
        [ADDR] SWF 1 at 0x8  - FWS Header
                [FILE] Carved SWF MD5: 2498e9c0701dc0e461ab4358f9102bc5.swf
```

### Example 2 - detecting and extracting a SWF file from a RTF document on Windows:

```text
C:\oletools>pyxswf -xf "rtf_flash.rtf"
RTF embedded object size 1498557 at index 000036DD
[SUMMARY] 1 SWF(s) in MD5:46a110548007e04f4043785ac4184558:RTF_embedded_object_0
00036DD
        [ADDR] SWF 1 at 0xc40  - FWS Header
                [FILE] Carved SWF MD5: 2498e9c0701dc0e461ab4358f9102bc5.swf
```

## How to use pyxswf in Python applications	

TODO

--------------------------------------------------------------------------

python-oletools documentation
-----------------------------

- [[Home]]
- [[License]]
- [[Install]]
- [[Contribute]], Suggest Improvements or Report Issues
- Tools:
	- [[mraptor]]
	- [[msodde]]
	- [[olebrowse]]
	- [[oledir]]
	- [[oleid]]
	- [[olemap]]
	- [[olemeta]]
	- [[oleobj]]
	- [[oletimes]]
	- [[olevba]]
	- [[pyxswf]]
	- [[rtfobj]]
