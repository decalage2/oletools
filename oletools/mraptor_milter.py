#!/usr/bin/env python
"""
mraptor_milter

mraptor_milter is a milter script for the Sendmail and Postfix e-mail
servers. It parses MS Office documents (e.g. Word, Excel) to detect
malicious macros. Documents with malicious macros are removed and
replaced by harmless text files.

Supported formats:
- Word 97-2003 (.doc, .dot), Word 2007+ (.docm, .dotm)
- Excel 97-2003 (.xls), Excel 2007+ (.xlsm, .xlsb)
- PowerPoint 97-2003 (.ppt), PowerPoint 2007+ (.pptm, .ppsm)
- Word 2003 XML (.xml)
- Word/Excel Single File Web Page / MHTML (.mht)
- Publisher (.pub)

Author: Philippe Lagadec - http://www.decalage.info
License: BSD, see source code or documentation

mraptor_milter is part of the python-oletools package:
http://www.decalage.info/python/oletools
"""

# === LICENSE ==================================================================

# mraptor_milter is copyright (c) 2016-2017 Philippe Lagadec (http://www.decalage.info)
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

# --- CHANGELOG --------------------------------------------------------------
# 2016-08-08 v0.01 PL: - first version
# 2016-08-12 v0.02 PL: - added logging to file with time rotation
#                      - archive each e-mail to a file before filtering
# 2016-08-30 v0.03 PL: - added daemonize to run as a Unix daemon
# 2016-09-06 v0.50 PL: - fixed issue #20, is_zipfile on Python 2.6
# 2017-04-26 v0.51 PL: - fixed absolute imports (issue #141)

__version__ = '0.51'

# --- TODO -------------------------------------------------------------------

# TODO: option to run in the foreground for troubleshooting
# TODO: option to write logs to the console
# TODO: options to set listening port and interface
# TODO: config file for all parameters
# TODO: option to run as a non-privileged user
# TODO: handle files in archives


# --- IMPORTS ----------------------------------------------------------------

import Milter         # not part of requirements, therefore: # pylint: disable=import-error
import io
import time
import email
import sys
import os
import logging
import logging.handlers
import datetime
import StringIO      # not part of requirements, therefore: # pylint: disable=import-error

from socket import AF_INET6

# IMPORTANT: it should be possible to run oletools directly as scripts
# in any directory without installing them with pip or setup.py.
# In that case, relative imports are NOT usable.
# And to enable Python 2+3 compatibility, we need to use absolute imports,
# so we add the oletools parent folder to sys.path (absolute+normalized path):
_thismodule_dir = os.path.normpath(os.path.abspath(os.path.dirname(__file__)))
# print('_thismodule_dir = %r' % _thismodule_dir)
_parent_dir = os.path.normpath(os.path.join(_thismodule_dir, '..'))
# print('_parent_dir = %r' % _thirdparty_dir)
if not _parent_dir in sys.path:
    sys.path.insert(0, _parent_dir)

from oletools import olevba, mraptor

from Milter.utils import parse_addr  # not part of requirements, therefore: # pylint: disable=import-error

from zipfile import is_zipfile



# --- CONSTANTS --------------------------------------------------------------

# TODO: read parameters from a config file
# at postfix smtpd_milters = inet:127.0.0.1:25252
SOCKET = "inet:25252@127.0.0.1"  # bind to unix or tcp socket "inet:port@ip" or "/<path>/<to>/<something>.sock"
TIMEOUT = 30  # Milter timeout in seconds
# CFG_DIR = "/etc/macromilter/"
# LOG_DIR = "/var/log/macromilter/"

# TODO: different path on Windows:
LOGFILE_DIR = '/var/log/mraptor_milter'
# LOGFILE_DIR = '.'
LOGFILE_NAME = 'mraptor_milter.log'
LOGFILE_PATH = os.path.join(LOGFILE_DIR, LOGFILE_NAME)

# Directory where to save a copy of each received e-mail:
ARCHIVE_DIR = '/var/log/mraptor_milter'
# ARCHIVE_DIR = '.'

# file to store PID for daemonize
PIDFILE = "/tmp/mraptor_milter.pid"



# === LOGGING ================================================================

# Set up a specific logger with our desired output level
log = logging.getLogger('MRMilter')

# disable logging by default - enable it in main app:
log.setLevel(logging.CRITICAL+1)

# NOTE: all logging config is done in the main app, not here.

# === CLASSES ================================================================

# Inspired from https://github.com/jmehnle/pymilter/blob/master/milter-template.py
# TODO: check https://github.com/sdgathman/pymilter which looks more recent

class MacroRaptorMilter(Milter.Base):
    '''
    '''
    def __init__(self):
        # A new instance with each new connection.
        # each connection runs in its own thread and has its own myMilter
        # instance.  Python code must be thread safe.  This is trivial if only stuff
        # in myMilter instances is referenced.
        self.id = Milter.uniqueID()  # Integer incremented with each call.
        self.message = None
        self.IP = None
        self.port = None
        self.flow = None
        self.scope = None
        self.IPname = None  # Name from a reverse IP lookup

    @Milter.noreply
    def connect(self, IPname, family, hostaddr):
        '''
        New connection (may contain several messages)
        :param IPname: Name from a reverse IP lookup
        :param family: IP version 4 (AF_INET) or 6 (AF_INET6)
        :param hostaddr: tuple (IP, port [, flow, scope])
        :return: Milter.CONTINUE
        '''
        # Examples:
        # (self, 'ip068.subnet71.example.com', AF_INET, ('215.183.71.68', 4720) )
        # (self, 'ip6.mxout.example.com', AF_INET6,
        #	('3ffe:80e8:d8::1', 4720, 1, 0) )
        self.IP = hostaddr[0]
        self.port = hostaddr[1]
        if family == AF_INET6:
            self.flow = hostaddr[2]
            self.scope = hostaddr[3]
        else:
            self.flow = None
            self.scope = None
        self.IPname = IPname  # Name from a reverse IP lookup
        self.message = None  # content
        log.info("[%d] connect from host %s at %s" % (self.id, IPname, hostaddr))
        return Milter.CONTINUE

    @Milter.noreply
    def envfrom(self, mailfrom, *rest):
        '''
        Mail From - Called at the beginning of each message within a connection
        :param mailfrom:
        :param str:
        :return: Milter.CONTINUE
        '''
        self.message = io.BytesIO()
        # NOTE: self.message is only an *internal* copy of message data.  You
        # must use addheader, chgheader, replacebody to change the message
        # on the MTA.
        self.canon_from = '@'.join(parse_addr(mailfrom))
        self.message.write('From %s %s\n' % (self.canon_from, time.ctime()))
        log.debug('[%d] Mail From %s %s\n' % (self.id, self.canon_from, time.ctime()))
        log.debug('[%d] mailfrom=%r, rest=%r' % (self.id, mailfrom, rest))
        return Milter.CONTINUE

    @Milter.noreply
    def envrcpt(self, to, *rest):
        '''
        RCPT TO
        :param to:
        :param str:
        :return: Milter.CONTINUE
        '''
        log.debug('[%d] RCPT TO %r, rest=%r\n' % (self.id, to, rest))
        return Milter.CONTINUE

    @Milter.noreply
    def header(self, header_field, header_field_value):
        '''
        Add header
        :param header_field:
        :param header_field_value:
        :return: Milter.CONTINUE
        '''
        self.message.write("%s: %s\n" % (header_field, header_field_value))
        return Milter.CONTINUE

    @Milter.noreply
    def eoh(self):
        '''
        End of headers
        :return: Milter.CONTINUE
        '''
        self.message.write("\n")
        return Milter.CONTINUE

    @Milter.noreply
    def body(self, chunk):
        '''
        Message body (chunked)
        :param chunk:
        :return: Milter.CONTINUE
        '''
        self.message.write(chunk)
        return Milter.CONTINUE

    def close(self):
        return Milter.CONTINUE

    def abort(self):
        '''
        Clean up if the connection is closed by client
        :return: Milter.CONTINUE
        '''
        return Milter.CONTINUE

    def archive_message(self):
        '''
        Save a copy of the current message in its original form to a file
        :return: nothing
        '''
        date_time = datetime.datetime.utcnow().isoformat('_')
        # assumption: by combining datetime + milter id, the filename should be unique:
        # (the only case for duplicates is when restarting the milter twice in less than a second)
        fname = 'mail_%s_%d.eml' % (date_time, self.id)
        fname = os.path.join(ARCHIVE_DIR, fname)
        log.debug('Saving a copy of the original message to file %r' % fname)
        open(fname, 'wb').write(self.message.getvalue())

    def eom(self):
        '''
        This method is called when the end of the email message has been reached.
        This event also triggers the milter specific actions
        :return: Milter.ACCEPT or Milter.DISCARD if processing error
        '''
        try:
            # set data pointer back to 0
            self.message.seek(0)
            self.archive_message()
            result = self.check_mraptor()
            if result is not None:
                return result
            else:
                return Milter.ACCEPT
                # if error make a fall-back to accept
        except Exception:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            log.exception("[%d] Unexpected error - fall back to ACCEPT: %s %s %s"
                              % (self.id, exc_type, fname, exc_tb.tb_lineno))
            return Milter.ACCEPT

    def check_mraptor(self):
        '''
        Check the attachments of a message using mraptor.
        If an attachment is identified as suspicious, it is replaced by a simple text file.
        :return: Milter.ACCEPT or Milter.DISCARD if processing error
        '''
        msg = email.message_from_string(self.message.getvalue())
        result = Milter.ACCEPT
        try:
            for part in msg.walk():
                # for name, value in part.items():
                #     log.debug(' - %s: %r' % (name, value))
                content_type = part.get_content_type()
                log.debug('[%d] Content-type: %r' % (self.id, content_type))
                # TODO: handle any content-type, but check the file magic?
                if not content_type.startswith('multipart'):
                    filename = part.get_filename(None)
                    log.debug('[%d] Analyzing attachment %r' % (self.id, filename))
                    attachment = part.get_payload(decode=True)
                    attachment_lowercase = attachment.lower()
                    # check if this is a supported file type (if not, just skip it)
                    # TODO: this function should be provided by olevba
                    if attachment.startswith(olevba.olefile.MAGIC) \
                        or is_zipfile(StringIO.StringIO(attachment)) \
                        or 'http://schemas.microsoft.com/office/word/2003/wordml' in attachment \
                        or ('mime' in attachment_lowercase and 'version' in attachment_lowercase
                            and 'multipart' in attachment_lowercase):
                        vba_parser = olevba.VBA_Parser(filename='message', data=attachment)
                        vba_code_all_modules = ''
                        for (subfilename, stream_path, vba_filename, vba_code) in vba_parser.extract_all_macros():
                            vba_code_all_modules += vba_code + '\n'
                        m = mraptor.MacroRaptor(vba_code_all_modules)
                        m.scan()
                        if m.suspicious:
                            log.warning('[%d] The attachment %r contains a suspicious macro: replace it with a text file'
                                            % (self.id, filename))
                            part.set_payload('This attachment has been removed because it contains a suspicious macro.')
                            part.set_type('text/plain')
                            # TODO: handle case when CTE is absent
                            part.replace_header('Content-Transfer-Encoding', '7bit')
                            # for name, value in part.items():
                            #     log.debug(' - %s: %r' % (name, value))
                            # TODO: archive filtered e-mail to a file
                        else:
                            log.debug('The attachment %r is clean.'
                                            % filename)
        except Exception:
            log.exception('[%d] Error while processing the message' % self.id)
            # TODO: depending on error, decide to forward the e-mail as-is or not
            result = Milter.DISCARD
        # TODO: only do this if the body has actually changed
        body = str(msg)
        self.message = io.BytesIO(body)
        self.replacebody(body)
        log.info('[%d] Message relayed' % self.id)
        return result


# === MAIN ===================================================================

def main():
    # banner
    print('mraptor_milter v%s - http://decalage.info/python/oletools' % __version__)
    print('logging to file %s' % LOGFILE_PATH)
    print('Press Ctrl+C to stop.')

    # make sure the log directory exists:
    try:
        os.makedirs(LOGFILE_DIR)
    except:
        pass
    # Add the log message handler to the logger
    # log to files rotating once a day:
    handler = logging.handlers.TimedRotatingFileHandler(LOGFILE_PATH, when='D', encoding='utf8')
    # create formatter and add it to the handlers
    formatter = logging.Formatter('%(asctime)s - %(levelname)8s: %(message)s')
    handler.setFormatter(formatter)
    log.addHandler(handler)
    # enable logging:
    log.setLevel(logging.DEBUG)

    log.info('Starting mraptor_milter v%s - listening on %s' % (__version__, SOCKET))
    log.debug('Python version: %s' % sys.version)

    # Register to have the Milter factory create instances of the class:
    Milter.factory = MacroRaptorMilter
    flags = Milter.CHGBODY + Milter.CHGHDRS + Milter.ADDHDRS
    flags += Milter.ADDRCPT
    flags += Milter.DELRCPT
    Milter.set_flags(flags)  # tell Sendmail which features we use
    # set the "last" fall back to ACCEPT if exception occur
    Milter.set_exception_policy(Milter.ACCEPT)
    # start the milter
    Milter.runmilter("mraptor_milter", SOCKET, TIMEOUT)
    log.info('Stopping mraptor_milter.')


if __name__ == "__main__":

    # Using daemonize:
    # See http://daemonize.readthedocs.io/en/latest/
    from daemonize import Daemonize    # not part of requirements, therefore: # pylint: disable=import-error
    daemon = Daemonize(app="mraptor_milter", pid=PIDFILE, action=main)
    daemon.start()

    # Using python-daemon - Does not work as-is, need to create the PID file
    # See https://pypi.org/project/python-daemon/
    # See PEP-3143: https://www.python.org/dev/peps/pep-3143/
    # import daemon
    # import lockfile
    # with daemon.DaemonContext(pidfile=lockfile.FileLock(PIDFILE)):
    #     main()
