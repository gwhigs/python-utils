import ftplib


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
