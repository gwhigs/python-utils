"""
A collection of utility functions for dealing with FTP servers.
"""
import ftplib
import os


class FTPUtilsException(Exception):
    pass


class ServerAddressMixin(object):
    """
    Mixin for ftplib's FTP class.

    Configures passive connections to use the originally supplied
    host address in the event the server reports an 'unroutable' address
    (i.e. starts with any of UNROUTABLE_IPS).

    Useful if your server is silly and reports back an internal IP
    during initial control, since ftplib uses the reported IP instead of
    the initially supplied host name for subsequent transfer commands
    in passive mode (active by default), most likely causing a timeout.

    May possibly be implemented in future Python releases as

    `passive_ignore_host = False`

    on the FTP and FTP_TLS classes themselves.
    """

    UNROUTABLE_IPS = ('192.168', '10.', '172.')

    def makepasv(self):
        host, port = super(ServerAddressMixin, self).makepasv()
        if any(host.startswith(s) for s in self.UNROUTABLE_IPS):
            host = self.host
        return host, port


class CustomFTP(ServerAddressMixin, ftplib.FTP):
    pass


class CustomFTP_TLS(ServerAddressMixin, ftplib.FTP_TLS):
    pass


def get_ftp(url, user, password, port=21, secure=False):
    """
    Convenience function for initializing FTP connections. 
    """
    ftp = CustomFTP_TLS() if secure else CustomFTP()
    ftp.connect(url, port)
    ftp.login(user, password)
    if secure:
        ftp.prot_p()
    return ftp


def upload_file_to_remote(bytes_file_obj, filename, target_ftp, target_path):
    """
    Stores a file on a remote FTP server.
    Note: file_obj must be opened in bytes mode.
    """
    if bytes_file_obj.mode != 'rb':
        raise FTPUtilsException('File must be opened in bytes mode.')
    cmd = 'STOR {}/{}'.format(target_path, filename)
    target_ftp.storbinary(cmd, bytes_file_obj)


def get_filenames(path):
    """
    Returns a list of all non-directory file names for a given path.
    """
    _path, _dirnames, filenames = next(os.walk(path))
    return filenames


def upload_all_to_remote(path, target_ftp, target_path):
    """
    Uploads all files in a given directory to a server via FTP.
    """
    files = get_filenames(path)
    for fn in files:
        with open(os.path.join(path, fn), 'rb') as f:
            upload_file_to_remote(f, fn, target_ftp, target_path)
            print('Report uploaded to server: {}'.format(fn))
