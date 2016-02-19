
import ConfigParser

class Test_Config:
    def __init__(self, config_file_path):
        cf = ConfigParser.ConfigParser()
        assert isinstance(config_file_path, object)
        cf.read(config_file_path)

        self.host = cf.get('BaseConf', 'host')
        self.caseDir = cf.get('BaseConf', 'caseDir')

        assert isinstance(self.caseDir, object)
        assert isinstance(self.host, object)
        print self.host, self.caseDir
