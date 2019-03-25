"""
Errors used in several tools to avoid duplication

.. codeauthor:: Intra2net AG <info@intra2net.com>
"""

class CryptoErrorBase(ValueError):
    """Base class for crypto-based exceptions."""
    pass


class CryptoLibNotImported(CryptoErrorBase, ImportError):
    """Exception thrown if msoffcrypto is needed but could not be imported."""

    def __init__(self):
        super(CryptoLibNotImported, self).__init__(
            'msoffcrypto-tools is not installed. Please run "pip install msoffcrypto-tool" or see https://github.com/nolze/msoffcrypto-tool')


class UnsupportedEncryptionError(CryptoErrorBase):
    """Exception thrown if file is encrypted and cannot deal with it."""
    def __init__(self, filename=None):
        super(UnsupportedEncryptionError, self).__init__(
            'Office file {}is encrypted, not yet supported'
            .format('' if filename is None else filename + ' '))


class WrongEncryptionPassword(CryptoErrorBase):
    """Exception thrown if encryption could be handled but passwords wrong."""
    def __init__(self, filename=None):
        super(WrongEncryptionPassword, self).__init__(
            'Given passwords could not decrypt office file{}, use option -p to specify the password'
            .format('' if filename is None else ' ' + filename))


class MaxCryptoNestingReached(CryptoErrorBase):
    """
    Exception thrown if decryption is too deeply layered.

    (...or decrypt code creates inf loop)
    """
    def __init__(self, n_layers, filename=None):
        super(MaxCryptoNestingReached, self).__init__(
            'Encountered more than {} layers of encryption for office file{}'
            .format(n_layers, '' if filename is None else ' ' + filename))
