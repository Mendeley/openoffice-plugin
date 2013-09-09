#!/usr/bin/python

import time
import sys
if sys.version_info < (3, 0):
    import httplib
else:
    import http.client as httplib
    StandardError = Exception

# Mendeley HTTP Client

# A client for communicating with the HTTP/JSON Mendeley Desktop Word
# processor API

# simplejson is json 
# simplejson is json (using a generic 'except' and not 'except ImportError'
# because MD-19770. See https://bugs.launchpad.net/ubuntu/+source/libreoffice/+bug/1222823
try: import simplejson as json
except: import json

# For communicating with the Mendeley Desktop HTTP API
class MendeleyHttpClient():
    HOST = "127.0.0.1" # much faster than "localhost" on Windows
                       # see http://cubicspot.blogspot.com/2010/07/fixing-slow-apache-on-localhost-under.html
    PORT = "50002"
    CONTENT_TYPE = "application/vnd.mendeley.wordProcessorApi+json; version=1.0"
    lastRequestTime = -1

    def __init__(self):
        self.previousResponse = self.Response(200, None, None, None)

    class Response:
        def __init__(self, status, contentType, body, request):
            self.status = status
            self.body = body
            self.contentType = contentType
            self.request = request

    class UnexpectedResponse(StandardError):
        def __init__(self, response):

            try:
                message = "response: ", json.dumps(response)
            except:
                message = "status: " + str(response.status)
                try:
                    # not sure why this works but the above doesn't
                    message += ", body: " + json.dumps(response.body.__dict__)
                except:
                    message += ", body: " + str(response.body)
            StandardError.__init__(self, message)

    # Currently this uses the same version number for all API routes,
    # this could be altered to be more fine-grained 
    class Request(object):    
        def __init__(self, verb, path, contentType, acceptType, body):
            self._verb = verb
            self._path = path
            self._body = body  # python dictionary
            self._contentType = contentType
            self._acceptType = acceptType
    
        def verb(self):
            return self._verb
    
        def path(self):
            return self._path
    
        def acceptType(self):
            return self._acceptType

        def contentType(self):
            return self._contentType

        def body(self):
            return json.dumps(self._body)

    class GetRequest(Request):
        def __init__(self, path):
            super(MendeleyHttpClient.GetRequest, self).__init__(
                "GET",
                path,
                "",
                MendeleyHttpClient.CONTENT_TYPE,
                "")
            
    class PostRequest(Request):
        def __init__(self, path, body):
            super(MendeleyHttpClient.PostRequest, self).__init__(
                "POST",
                path,
                MendeleyHttpClient.CONTENT_TYPE,
                MendeleyHttpClient.CONTENT_TYPE,
                body)

    def formattedCitationsAndBibliography_Interactive(self, citationStyleUrl, citationClusters):
        request = self.PostRequest(
            "/formattedCitationsAndBibliography/interactive",
            {
                "citationStyleUrl": citationStyleUrl,
                "citationClusters": citationClusters
            }
            )
        return self.request(request)
    
    def citation_choose_interactive(self, citationEditorHint):
        request = self.PostRequest(
            "/citation/choose/interactive",
            citationEditorHint
            )
        return self.request(request)

    def citation_edit_interactive(self, citationCluster):
        request = self.PostRequest(
            "/citation/edit/interactive",
            citationCluster
            )
        return self.request(request)

    def citation_update_interactive(self, formattedCitationCluster):
        request = self.PostRequest(
            "/citation/update/interactive",
            formattedCitationCluster
            )
        return self.request(request)

    def citationStyle_choose_interactive(self, currentStyleUrl):
        request = self.PostRequest(
            "/citationStyle/choose/interactive",
            currentStyleUrl
            )
        return self.request(request)

    def styleName_getFromUrl(self, styleUrl):
        request = self.PostRequest(
            "/citationStyle/getNameFromUrl",
            styleUrl
            )
        return self.request(request)

    def citationStyles_default(self):
        request = self.GetRequest(
            "/citationStyles/default",
            )
        return self.request(request)

    def citations_merge(self, citationClusters):
        request = self.PostRequest(
            "/citations/merge",
            citationClusters
            )
        return self.request(request)

    def citation_undoManualFormat(self, citationCluster):
        request = self.PostRequest(
            "/citation/undoManualFormat",
            citationCluster
            )
        return self.request(request)

    def wordProcessor_set(self, wordProcessor):
        request = self.PostRequest(
            "/wordProcessor/set",
            wordProcessor
            )
        return self.request(request)

    def testMethods_citationCluster_getFromUuid(self, uuid):
        request = self.PostRequest(
            "/testMethods/citationCluster/getFromUuid",
            uuid
            )
        return self.request(request)
        
    def userAccount(self):
        request = self.GetRequest(
            "/userAccount"
            )
        return self.request(request)

    def mendeleyDesktopInfo(self):
        request = self.GetRequest(
            "/mendeleyDesktopInfo"
            )
        return self.request(request)

    # Need to define a class for this.
    # I tried using a object() instance but it doesn't contain a __dict__
    class ResponseBody:
        pass

    # Sets up a connection to Mendeley Desktop, makes a HTTP request and
    # returns the data
    def request(self, requestData):
        headers = {}
        # putting an empty string in Content-Type causes the Mendeley Desktop HTTP
        # server to put the Accept header value in a field called "content-typeaccept"
        # TODO: check where this error comes from
        if requestData.contentType() != "":
            headers["Content-Type"] = requestData.contentType()
        if requestData.acceptType() != "":
            headers["Accept"] = requestData.acceptType()
        startTime = time.time()
        connection = httplib.HTTPConnection(self.HOST + ":" + self.PORT)
        connection.request(requestData.verb(), requestData.path(), requestData.body(), headers)
        response = connection.getresponse()
        data = response.read()
        data = data.decode('utf-8')

        if (response.status == 200 and (not response.getheader("Content-Type") is None) and
                response.getheader("Content-Type") != requestData.acceptType()):
            # TODO: abort if the wrong content type is returned
            pass

        responseBody = MendeleyHttpClient.ResponseBody()
        try:
            responseBody.__dict__.update(json.loads(data))
        except:
            responseBody = data
        connection.close()
        self.lastRequestTime = 1000 * (time.time() - startTime)

        self.previousResponse = \
                self.Response(response.status, response.getheader("Content-Type"), responseBody,
                        requestData)

        return self.previousResponse
