"""
Errors used in several tools to avoid duplication

.. codeauthor:: Intra2net AG <info@intra2net.com>
"""

class FileIsEncryptedError(ValueError):
    """Exception thrown if file is encrypted and cannot deal with it."""
    # see also: same class in olevba[3] and record_base
    def __init__(self, filename):
        super(FileIsEncryptedError, self).__init__(
            'Office file {} is encrypted, not yet supported'.format(filename))
