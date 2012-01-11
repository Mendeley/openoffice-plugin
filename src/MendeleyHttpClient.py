#!/usr/bin/python

import httplib
import time

# Mendeley HTTP Client

# A client for communicating with the HTTP/JSON Mendeley Desktop Word
# processor API

# simplejson is json 
try: import simplejson as json
except ImportError: import json

# For communicating with the Mendeley Desktop HTTP API
class MendeleyHttpClient():
    HOST = "127.0.0.1" # much faster than "localhost" on Windows
                       # see http://cubicspot.blogspot.com/2010/07/fixing-slow-apache-on-localhost-under.html
    PORT = "5000"
    API_VERSION = "1.0"
    lastRequestTime = -1

    def __init__(self):
        pass

    class UnexpectedResponse(StandardError):
        def __init__(self, response):
            try:
                message = json.dumps(response)
            except:
                message = "status: " + str(response.status)
                try:
                    # not sure why this works but the above doesn't
                    message += ", body: " + json.dumps(response.body.__dict__)
                except:
                    pass
            StandardError.__init__(self, message)

    class Request(object):    
        def __init__(self, verb, path, contentType, acceptType, body):
            self._verb = verb
            self._path = path
            self._contentType = contentType
            self._acceptType = acceptType
            self._body = body  # python dictionary
            self._versionSuffix = ";version=" + MendeleyHttpClient.API_VERSION
            self._rootPath = "mendeley/wordProcessorApi"
    
        def verb(self):
            return self._verb
    
        def path(self):
            return self._path
    
        def acceptType(self):
            return self._acceptType + self._versionSuffix

        def contentType(self):
            return self._contentType + self._versionSuffix

        def body(self):
            return json.dumps(self._body)

    class GetRequest(Request):
        def __init__(self, path, acceptType):
            super(MendeleyHttpClient.GetRequest, self).__init__(
                "GET",
                path,
                "",
                acceptType,
                "")
            
    class PostRequest(Request):
        def __init__(self, path, contentType, acceptType, body):
            super(MendeleyHttpClient.PostRequest, self).__init__(
                "POST",
                path,
                contentType,
                acceptType,
                body)

    class PutRequest(Request):
        def __init__(self, path, contentType, body):
            super(MendeleyHttpClient.PutRequest, self).__init__(
                "PUT",
                path,
                contentType,
                "",
                "")

    def formattedCitationsAndBibliography_Interactive(self, citationStyleUrl, citationClusters):
        request = self.PostRequest(
            "/formattedCitationsAndBibliography/interactive",
            "mendeley/wordProcessorApi/documentToFormat+json",
            "mendeley/wordProcessorApi/formattedDocument+json",
            {
                "citationStyleUrl": citationStyleUrl,
                "citationClusters": citationClusters
            }
            )
        return self.request(request)
    
    def citation_choose_interactive(self):
        request = self.GetRequest(
            "/citation/choose/interactive",
            "mendeley/citationStyleUrl+json"
            )
        return self.request(request)

    def citation_edit_interactive(self, citationCluster):
        request = self.PostRequest(
            "/citation/edit/interactive",
            "mendeley/citationCluster+json",
            "mendeley/editedCitationCluster+json",
            citationCluster
            )
        return self.request(request)

    def citation_update_interactive(self, formattedCitationCluster):
        request = self.PostRequest(
            "/citation/update/interactive",
            "mendeley/formattedCitationCluster+json",
            "mendeley/editedCitationCluster+json",
            formattedCitationCluster
            )
        return self.request(request)

    def citationStyle_choose_interactive(self, currentStyleUrl):
        request = self.PostRequest(
            "/citationStyle/choose/interactive",
            "mendeley/citationStyleUrl+json",
            "mendeley/citationStyleUrl+json",
            currentStyleUrl
            )
        return self.request(request)

    def styleName_getFromUrl(self, styleUrl):
        request = self.PostRequest(
            "/citationStyle/getNameFromUrl",
            "mendeley/citationStyleUrl+json",
            "mendeley/citationStyleName+json",
            styleUrl
            )
        return self.request(request)

    def bringPluginToForeground(self):
        request = self.GetRequest(
            "/bringPluginToForeground",
            "mendeley/bringToForegroundSuccess+json"
            )
        return self.request(request)

    def citationStyles_default(self):
        request = self.GetRequest(
            "/citationStyles/default",
            "mendeley/citationStyles+json"
            )
        return self.request(request)

    def citations_merge(self, citationClusters):
        request = self.PostRequest(
            "/citations/merge",
            "mendeley/citationClusters+json",
            "mendeley/citationCluster+json",
            citationClusters
            )
        return self.request(request)

    def citation_undoManualFormat(self, citationCluster):
        request = self.PostRequest(
            "/citation/undoManualFormat",
            "mendeley/citationCluster+json",
            "mendeley/citationCluster+json",
            citationCluster
            )
        return self.request(request)

    def wordProcessor_set(self, wordProcessor):
        request = self.PostRequest(
            "/wordProcessor/set",
            "mendeley/wordProcessor+json",
            "",
            wordProcessor
            )
        return self.request(request)

    def testMethods_citationCluster_getFromUuid(self, uuid):
        request = self.PostRequest(
            "/testMethods/citationCluster/getFromUuid",
            "mendeley/referenceUuid+json",
            "mendeley/citationCluster+json",
            uuid
            )
        return self.request(request)
        
    def userAccount(self):
        request = self.GetRequest(
            "/userAccount",
            "mendeley/getUserAccount+json")
        return self.request(request)

    # Need to define a class for this.
    # I tried using a object() instance but it doesn't contain a __dict__
    class ResponseBody:
        pass

    class Response:
        def __init__(self, status, body):
            self.status = status
            self.body = body

    # Sets up a connection to Mendeley Desktop, makes a HTTP request and
    # returns the data
    def request(self, requestData):
        headers = { "Content-Type" : requestData.contentType(), "Accept" : requestData.acceptType() }
        startTime = time.time()
        connection = httplib.HTTPConnection(self.HOST + ":" + self.PORT)
        connection.request(requestData.verb(), requestData.path(), requestData.body(), headers)
        response = connection.getresponse()
        data = response.read()
        data = data.decode('utf-8')

        if response.getheader("Content-Type") != requestData.acceptType():
            # TODO: abort if the wrong content type is returned
            #print "WARNING: server returned wrong content-type"
            #return
            pass
            with open("f:\MendeleyHttpClient.log", "a") as logFile:
                logFile.write("WARNING: server returned wrong content-type\n")
        
        responseBody = MendeleyHttpClient.ResponseBody()
        #print "data = " + data
        try:
            responseBody.__dict__.update(json.loads(data))
        except:
            responseBody = data
        connection.close()
        self.lastRequestTime = 1000 * (time.time() - startTime)
        return self.Response(response.status, responseBody)

