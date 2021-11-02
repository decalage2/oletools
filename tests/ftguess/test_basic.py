"""Test ftguess"""

import unittest
import os
from os.path import splitext
from oletools import ftguess

# Directory with test data, independent of current working directory
from tests.test_utils import DATA_BASE_DIR
from tests.test_utils.testdata_reader import loop_over_files


class TestFTGuess(unittest.TestCase):
    """Test ftguess"""

    def test_all(self):
        """Run all files in test-data and compare to known ouput"""
        # ftguess knows extension for each FType, create a reverse mapping
        used_types = (
            ftguess.FType_RTF, ftguess.FType_Generic_OLE,
            ftguess.FType_Generic_Zip, ftguess.FType_Word97,
            ftguess.FType_Word2007, ftguess.FType_Word2007_Macro,
            ftguess.FType_Word2007_Template,
            ftguess.FType_Word2007_Template_Macro, ftguess.FType_Excel97,
            ftguess.FType_Excel2007,
            ftguess.FType_Excel2007_XLSX , ftguess.FType_Excel2007_XLSM ,
            ftguess.FType_Excel2007_Template,
            ftguess.FType_Excel2007_Template_Macro,
            ftguess.FType_Excel2007_Addin_Macro, ftguess.FType_Powerpoint97,
            ftguess.FType_Powerpoint2007_Presentation,
            ftguess.FType_Powerpoint2007_Slideshow,
            ftguess.FType_Powerpoint2007_Macro,
            ftguess.FType_Powerpoint2007_Slideshow_Macro,
            ftguess.FType_XPS,
        )
        ftype_for_extension = dict()
        for ftype in used_types:
            for extension in ftype.extensions:
                ftype_for_extension[extension] = ftype

        # TODO: xlsb is not implemented yet
        ftype_for_extension['xlsb'] = ftguess.FType_Generic_OpenXML

        for filename, file_contents in loop_over_files():
            # let the system guess
            guess = ftguess.ftype_guess(data=file_contents)
            #print(f'for debugging: {filename} --> {guess}')

            # determine what we expect...
            before_dot, extension = splitext(filename)
            if extension == '.zip':
                extension = splitext(before_dot)[1]
            elif filename in ('basic/empty', 'basic/text'):
                extension = '.csv'    # have just like that
            elif not extension:
                self.fail('Could not find extension for test sample {0}'
                          .format(filename))
            extension = extension[1:]      # remove the leading '.'

            # encrypted files are mostly recognized (yet?), except .xls
            if filename.startswith('encrypted/'):
                if extension == 'xls':
                    expect = ftguess.FType_Excel97
                else:
                    expect = ftguess.FType_Generic_OLE

            elif extension in ('xml', 'csv', 'odt', 'ods', 'odp', 'potx', 'potm'):
                # not really an office file type
                expect = ftguess.FType_Unknown

            elif filename == 'basic/encrypted.docx':
                expect = ftguess.FType_Generic_OLE

            else:
                # other files behave nicely, so extension determines the type
                expect = ftype_for_extension[extension]

                self.assertEqual(guess.container, expect.container,
                                 msg='ftguess guessed container {0} for {1} '
                                     'but we expected {2}'
                                     .format(guess.container, filename,
                                             expect.container))
                self.assertEqual(guess.filetype, expect.filetype,
                                 msg='ftguess guessed filetype {0} for {1} '
                                     'but we expected {2}'
                                     .format(guess.filetype, filename,
                                             expect.filetype))
                self.assertEqual(guess.application, expect.application,
                                 msg='ftguess guessed application {0} for {1} '
                                     'but we expected {2}'
                                     .format(guess.application, filename,
                                             expect.application))

            if expect not in (ftguess.FType_Generic_OLE, ftguess.FType_Unknown):
                self.assertEqual(guess.is_excel(), extension.startswith('x')
                                                   and extension != 'xml'
                                                   and extension != 'xlsb'
                                                   and extension != 'xps')
                   # xlsb is excel but not handled properly yet
                self.assertEqual(guess.is_word(), extension.startswith('d'))
                self.assertEqual(guess.is_powerpoint(),
                                 extension.startswith('p'))



# just in case somebody calls this file as a script
if __name__ == '__main__':
    unittest.main()
