import pysftp
cnopts = pysftp.CnOpts()
cnopts.hostkeys = None
import os


class SFTP:

    def __init__(self, localpath, remotepath):
        self.cnopts = pysftp.CnOpts()
        self.cnopts.hostkeys = None
        self.localpath = localpath
        self.remotepath = remotepath
        
        self.logs = [os.environ.get('FTP_HOST'),os.environ.get('FTP_USER'),os.environ.get('FTP_PASSWORD'),22222]

    def send_file(self):

        sftp = pysftp.Connection(self.logs[0], username=self.logs[1], password=self.logs[2], port=self.logs[3], cnopts=self.cnopts)

        sftp.put(self.localpath, self.remotepath)

        sftp.close()




