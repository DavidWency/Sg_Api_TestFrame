import httplib, time, hashlib, urllib

class TestFrameLib:
    def __init__(self, sFile, IPPort):
        self.sFile = sFile
        self.IPPort = IPPort
        self.httpClient = None
        self.headers = {'Content-type': 'application/x-www-form-urlencoded', 'Accept': 'text/plain'}
        self.params = None
        self.timestamp = int(time.time())
        pass

    def HTTPInvoke(self, url):
        try:
            self.httpClient = httplib.HTTPConnection(self.IPPort)
            self.httpClient.request('POST', url)

            response = self.httpClient.getresponse()
            data = response.read()
            assert isinstance(data, object)
            return data
        except Exception, e:
            print e
        finally:
            if self.httpClient:
                self.httpClient.close()

    def HTTPRequest(self, url, param, method='POST', header=None):
        try:
            if method == 'GET':
                self.httpClient = httplib.HTTPConnection(self.IPPort)
                self.httpClient.request(method, url)
            else:
                self.httpClient = httplib.HTTPConnection(self.IPPort)
                self.params = urllib.urlencode(param)
                if header:
                    self.headers = header
                self.httpClient.request(method, url, self.params, self.headers)

            response = self.httpClient.getresponse()
            data = response.read()
            assert isinstance(data, object)
            return data
        except Exception, e:
            print e
        finally:
            if self.httpClient:
                self.httpClient.close()

    def get_result_code(self,result):
        position = result.find(',')
        if position != -1:
            result_code =  result[8: position]
        else:
            result_code = 'case is bad'
        return result_code

    def md5_src(src):
        m = hashlib.md5()
        m.update(src)
        sign = m.hexdigest()
        return sign
