""" Basic tests for ooxml.py """

import unittest

import os
from os.path import join, splitext
from tests.test_utils import DATA_BASE_DIR
from olefile import isOleFile
from oletools import ooxml
import logging


class TestOOXML(unittest.TestCase):
    """ Tests correct behaviour of XML parser """

    DO_DEBUG = False

    def setUp(self):
        if self.DO_DEBUG:
            logging.basicConfig(level=logging.DEBUG)


    def test_rough_doctype(self):
        """Checks all samples, expect either ole files or good ooxml output"""
        # map from extension to expected doctype
        ext2doc = dict(
            docx=ooxml.DOCTYPE_WORD, docm=ooxml.DOCTYPE_WORD,
            dotx=ooxml.DOCTYPE_WORD, dotm=ooxml.DOCTYPE_WORD,
            xml=(ooxml.DOCTYPE_EXCEL_XML, ooxml.DOCTYPE_WORD_XML),
            xlsx=ooxml.DOCTYPE_EXCEL, xlsm=ooxml.DOCTYPE_EXCEL,
            xlsb=ooxml.DOCTYPE_EXCEL, xlam=ooxml.DOCTYPE_EXCEL,
            xltx=ooxml.DOCTYPE_EXCEL, xltm=ooxml.DOCTYPE_EXCEL,
            pptx=ooxml.DOCTYPE_POWERPOINT, pptm=ooxml.DOCTYPE_POWERPOINT,
            ppsx=ooxml.DOCTYPE_POWERPOINT, ppsm=ooxml.DOCTYPE_POWERPOINT,
            potx=ooxml.DOCTYPE_POWERPOINT, potm=ooxml.DOCTYPE_POWERPOINT,
            ods=ooxml.DOCTYPE_NONE, odt=ooxml.DOCTYPE_NONE,
            odp=ooxml.DOCTYPE_NONE,
        )

        # files that are neither OLE nor xml:
        except_files = 'empty', 'text'
        except_extns = 'rtf', 'csv', 'zip'

        # analyse all files in data dir
        for base_dir, _, files in os.walk(DATA_BASE_DIR):
            for filename in files:
                if filename in except_files:
                    if self.DO_DEBUG:
                        print('skip file: ' + filename)
                    continue
                extn = splitext(filename)[1]
                if extn:
                    extn = extn[1:]      # remove the dot
                if extn in except_extns:
                    if self.DO_DEBUG:
                        print('skip extn: ' + filename)
                    continue

                full_name = join(base_dir, filename)
                if isOleFile(full_name):
                    if self.DO_DEBUG:
                        print('skip ole: ' + filename)
                    continue
                acceptable = ext2doc[extn]
                if not isinstance(acceptable, tuple):
                    acceptable = (acceptable, )
                try:
                    doctype = ooxml.get_type(full_name)
                except Exception:
                    self.fail('Failed to get doctype of {0}'.format(filename))
                self.assertTrue(doctype in acceptable,
                                msg='Doctype "{0}" for {1} not acceptable'
                                    .format(doctype, full_name))
                if self.DO_DEBUG:
                    print('ok: {0} --> {1}'.format(filename, doctype))

    def test_iter_all(self):
        """ test iter_xml without args """
        expect_subfiles = dict([
            ('[Content_Types].xml', 11),
            ('_rels/.rels', 4),
            ('word/_rels/document.xml.rels', 6),
            ('word/document.xml', 102),
            ('word/theme/theme1.xml', 227),
            ('word/settings.xml', 40),
            ('word/fontTable.xml', 25),
            ('word/webSettings.xml', 3),
            ('docProps/app.xml', 26),
            ('docProps/core.xml', 10),
            ('word/styles.xml', 441),
        ])
        n_elems = 0
        testfile = join(DATA_BASE_DIR, 'msodde', 'harmless-clean.docx')
        for subfile, elem, depth in ooxml.XmlParser(testfile).iter_xml():
            n_elems += 1
            if depth > 0:
                continue

            # now depth == 0; should occur once at end of every subfile
            if subfile not in expect_subfiles:
                self.fail('Subfile {0} not expected'.format(subfile))
            self.assertEqual(n_elems, expect_subfiles[subfile],
                             'wrong number of elems ({0}) yielded from {1}'
                             .format(n_elems, subfile))
            _ = expect_subfiles.pop(subfile)
            n_elems = 0

        self.assertEqual(len(expect_subfiles), 0,
                         'Forgot to iterate through subfile(s) {0}'
                         .format(expect_subfiles.keys()))

    def test_iter_subfiles(self):
        """ test that limitation on few subfiles works """
        testfile = join(DATA_BASE_DIR, 'msodde', 'dde-test.xlsx')
        subfiles = ['xl/theme/theme1.xml', 'docProps/app.xml']
        parser = ooxml.XmlParser(testfile)
        for subfile, elem, depth in parser.iter_xml(subfiles):
            if self.DO_DEBUG:
                print(u'{0} {1}{2}'.format(subfile, '  '*depth,
                                           ooxml.debug_str(elem)))
            if subfile not in subfiles:
                self.fail('should have been skipped: {0}'.format(subfile))
            if depth == 0:
                subfiles.remove(subfile)

        self.assertEqual(subfiles, [], 'missed subfile(s) {0}'
                                       .format(subfiles))

    def test_iter_tags(self):
        """ test that limitation to tags works """
        testfile = join(DATA_BASE_DIR, 'msodde', 'harmless-clean.docm')
        nmspc = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        tag = '{' + nmspc + '}p'

        parser = ooxml.XmlParser(testfile)
        n_found = 0
        for subfile, elem, depth in parser.iter_xml(tags=tag):
            n_found += 1
            self.assertEqual(elem.tag, tag)

            # also check that children are present
            n_children = 0
            for child in elem:
                n_children += 1
                self.assertFalse(child.tag == '')
            self.assertTrue(n_children > 0, 'no children for elem {0}'
                                            .format(ooxml.debug_str(elem)))

        self.assertEqual(n_found, 7)


# just in case somebody calls this file as a script
if __name__ == '__main__':
    unittest.main()
