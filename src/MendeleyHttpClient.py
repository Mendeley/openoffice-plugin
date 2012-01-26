#!/usr/bin/python

import httplib
import time
import sys

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
            if self._acceptType == "":
                return ""
            else:
                return self._acceptType + self._versionSuffix

        def contentType(self):
            if self._contentType == "":
                return ""
            else:
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
            "application/vnd.mendeley.documentToFormat+json",
            "application/vnd.mendeley.formattedDocument+json",
            {
                "citationStyleUrl": citationStyleUrl,
                "citationClusters": citationClusters
            }
            )
        return self.request(request)
    
    def citation_choose_interactive(self, citationEditorHint):
        request = self.PostRequest(
            "/citation/choose/interactive",
            "application/vnd.mendeley.citationEditorHint+json",
            "application/vnd.mendeley.editedCitationCluster+json",
            citationEditorHint
            )
        return self.request(request)

    def citation_edit_interactive(self, citationCluster):
        request = self.PostRequest(
            "/citation/edit/interactive",
            "application/vnd.mendeley.citationClusterWithHint+json",
            "application/vnd.mendeley.editedCitationCluster+json",
            citationCluster
            )
        return self.request(request)

    def citation_update_interactive(self, formattedCitationCluster):
        request = self.PostRequest(
            "/citation/update/interactive",
            "application/vnd.mendeley.formattedCitationCluster+json",
            "application/vnd.mendeley.editedCitationCluster+json",
            formattedCitationCluster
            )
        return self.request(request)

    def citationStyle_choose_interactive(self, currentStyleUrl):
        request = self.PostRequest(
            "/citationStyle/choose/interactive",
            "application/vnd.mendeley.citationStyleUrl+json",
            "application/vnd.mendeley.citationStyleUrl+json",
            currentStyleUrl
            )
        return self.request(request)

    def styleName_getFromUrl(self, styleUrl):
        request = self.PostRequest(
            "/citationStyle/getNameFromUrl",
            "application/vnd.mendeley.citationStyleUrl+json",
            "application/vnd.mendeley.citationStyleName+json",
            styleUrl
            )
        return self.request(request)

    def bringPluginToForeground(self):
        request = self.GetRequest(
            "/bringPluginToForeground",
            "application/vnd.mendeley.bringToForegroundResponse+json"
            )
        return self.request(request)

    def citationStyles_default(self):
        request = self.GetRequest(
            "/citationStyles/default",
            "application/vnd.mendeley.citationStyles+json"
            )
        return self.request(request)

    def citations_merge(self, citationClusters):
        request = self.PostRequest(
            "/citations/merge",
            "application/vnd.mendeley.citationClusters+json",
            "application/vnd.mendeley.citationCluster+json",
            citationClusters
            )
        return self.request(request)

    def citation_undoManualFormat(self, citationCluster):
        request = self.PostRequest(
            "/citation/undoManualFormat",
            "application/vnd.mendeley.citationCluster+json",
            "application/vnd.mendeley.citationCluster+json",
            citationCluster
            )
        return self.request(request)

    def wordProcessor_set(self, wordProcessor):
        request = self.PostRequest(
            "/wordProcessor/set",
            "application/vnd.mendeley.wordProcessor+json",
            "",
            wordProcessor
            )
        return self.request(request)

    def testMethods_citationCluster_getFromUuid(self, uuid):
        request = self.PostRequest(
            "/testMethods/citationCluster/getFromUuid",
            "application/vnd.mendeley.referenceUuid+json",
            "application/vnd.mendeley.citationCluster+json",
            uuid
            )
        return self.request(request)
    
    def testMethods_citationCluster_getFromUuid_deprecatedResponse(self, uuid):
        request = self.PostRequest(
            "/testMethods/citationCluster/getFromUuid",
            "application/vnd.mendeley.referenceUuid+json",
            "application/vnd.mendeley.deprecatedResponse+json",
            uuid
            )
        return self.request(request)

    def testMethods_citationCluster_getFromUuid_unknownResponse(self, uuid):
        request = self.PostRequest(
            "/testMethods/citationCluster/getFromUuid",
            "application/vnd.mendeley.referenceUuid+json",
            "application/vnd.mendeley.unknownResponse+json",
            uuid
            )
        return self.request(request)
        
    def userAccount(self):
        request = self.GetRequest(
            "/userAccount",
            "application/vnd.mendeley.userAccount+json")
        return self.request(request)

    def mendeleyDesktopVersion(self):
        request = self.GetRequest(
            "/mendeleyDesktopVersion",
            "application/vnd.mendeley.mendeleyDesktopVersion+json")
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
            sys.stderr.write("ERROR: server returned wrong content-type: " + response.getheader("Content-Type"))
            #return

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
