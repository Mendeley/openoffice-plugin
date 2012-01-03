#!/usr/bin/python

import httplib

# Mendeley HTTP Client

# A client for communicating with the HTTP/JSON Mendeley Desktop Word
# processor API

# to run test from command prompt:
# open Mendeley Desktop with testDatabase@test.com@local.sqlite
# python -c "from MendeleyHttpClient import test; test()"

# simplejson is json 
try: import simplejson as json
except ImportError: import json

# For communicating with the Mendeley Desktop HTTP API
class MendeleyHttpClient():
    HOST = "127.0.0.1" # much faster than "localhost"
    PORT = "5001"
    API_VERSION = "1.0"

    def __init__(self):
        pass

    class Request(object):    
        def __init__(self, verb, path, contentType, body):
            self._verb = verb
            self._path = path
            self._contentType = contentType
            self._body = body  # python dictionary
    
        def verb(self):
            return self._verb
    
        def path(self):
            return self._path
    
        def contentType(self):
            return self._contentType + ";version=" + MendeleyHttpClient.API_VERSION
    
        def body(self):
            return json.dumps(self._body)

    class FormattedCitationsAndBibliographyResponse:
        def __init__(self):
            self._contentType = "mendeley/formattedCitationsAndBibliography+json"
            self.citationStyleUrl = ""
            self.citationClusters = []
            self.bibliography = ""

        def contentType(self):
            return self._contentType + ";version=" + MendeleyHttpClient.API_VERSION

    def formattedCitationsAndBibliography_Interactive(self, citationStyleUrl, citationClusters):
        httpRequest = MendeleyHttpClient.Request(
            "POST",
            "/formattedCitationsAndBibliography/interactive",
            "mendeley/wordProcessorDocument+json",
            {
                "citationStyleUrl": citationStyleUrl,
                "citationClusters": citationClusters
            }
            )

        response = MendeleyHttpClient.FormattedCitationsAndBibliographyResponse()
        self.request(httpRequest, response)
        return response
        
    # Sets up a connection to Mendeley Desktop, makes a HTTP request and
    # returns the data
    def request(self, requestData, responseData):
        headers = { "Content-Type" : requestData.contentType(), "Accept" : responseData.contentType() }
        connection = httplib.HTTPConnection(self.HOST + ":" + self.PORT)
        connection.request(requestData.verb(), requestData.path(), requestData.body(), headers)
        response = connection.getresponse()
        data = response.read()
        data = data.decode('utf-8')

        # unescape quotation marks
        # TODO: check if it's correct that the server escapes quotation marks:
        data = data.replace('\\"', '"')

        # remove quotation mark from start and end of string
        # TODO: why are these present?
        #if data[0] == '"' and data[len(data) - 1] == '"':
        #    data = data[1:len(data) - 1]
        
        print "data: " + data

        print "response Content-Type = " + response.getheader("Content-Type")
        if response.getheader("Content-Type") != responseData.contentType():
			# TODO: abort if the wrong content type is returned
            print "WARNING: server returned wrong content-type"
            #return
        
        responseData.__dict__.update(json.loads(data))
        connection.close()
        return

def test():
    client = MendeleyHttpClient()

    # this call should complete without requiring user interaction
    response = client.formattedCitationsAndBibliography_Interactive(
            "http://www.zotero.org/styles/apa",
            [
                {
                    "citationCluster": json.loads("{ \"citationItems\" : [ { \"id\" : \"ITEM-1\", \"itemData\" : { \"author\" : [ { \"family\" : \"Smith\", \"given\" : \"John\" }, { \"family\" : \"Jr\", \"given\" : \"John Smith\" } ], \"id\" : \"ITEM-1\", \"issued\" : { \"date-parts\" : [ [ \"2001\" ] ] }, \"title\" : \"Title01\", \"type\" : \"article\" }, \"uris\" : [ \"http://local/documents/?uuid=55ff8735-3f3c-4c9f-87c3-8db322ba3f74\" ] }, { \"id\" : \"ITEM-2\", \"itemData\" : { \"author\" : [ { \"family\" : \"Evans\", \"given\" : \"Gareth\" }, { \"family\" : \"Jr\", \"given\" : \"Gareth Evans\" } ], \"id\" : \"ITEM-2\", \"issued\" : { \"date-parts\" : [ [ \"2002\" ] ] }, \"title\" : \"Title02\", \"type\" : \"article\" }, \"uris\" : [ \"http://local/documents/?uuid=15d6d1e4-a9ff-4258-88b6-a6d6d6bdc0ed\" ] } ], \"mendeley\" : { \"previouslyFormattedCitation\" : \"(Evans & Jr, 2002; Smith & Jr, 2001)\" }, \"properties\" : { \"noteIndex\" : 0 }, \"schema\" : \"https://github.com/citation-style-language/schema/raw/master/csl-citation.json\" }"),
                    "formattedText": "(Evans & Jr, 2002; Smith & Jr, 2001)"
                },
                {
                    "citationCluster": json.loads("{ \"citationItems\" : [ { \"id\" : \"ITEM-1\", \"itemData\" : { \"author\" : [ { \"family\" : \"Smith\", \"given\" : \"John\" }, { \"family\" : \"Jr\", \"given\" : \"John Smith\" } ], \"id\" : \"ITEM-1\", \"issued\" : { \"date-parts\" : [ [ \"2001\" ] ] }, \"title\" : \"Title01\", \"type\" : \"article\" }, \"uris\" : [ \"http://local/documents/?uuid=55ff8735-3f3c-4c9f-87c3-8db322ba3f74\" ] }, { \"id\" : \"ITEM-2\", \"itemData\" : { \"author\" : [ { \"family\" : \"Evans\", \"given\" : \"Gareth\" }, { \"family\" : \"Jr\", \"given\" : \"Gareth Evans\" } ], \"id\" : \"ITEM-2\", \"issued\" : { \"date-parts\" : [ [ \"2002\" ] ] }, \"title\" : \"Title02\", \"type\" : \"article\" }, \"uris\" : [ \"http://local/documents/?uuid=15d6d1e4-a9ff-4258-88b6-a6d6d6bdc0ed\" ] } ], \"mendeley\" : { \"previouslyFormattedCitation\" : \"(Evans & Jr, 2002; Smith & Jr, 2001)\" }, \"properties\" : { \"noteIndex\" : 0 }, \"schema\" : \"https://github.com/citation-style-language/schema/raw/master/csl-citation.json\" }"),
                    "formattedText": "(Evans & Jr, 2002; Smith & Jr, 2001)"
                }
            ]
            )
    print "response: " + json.dumps(response.__dict__)

    for cluster in response.citationClusters:
        print "cluster: " + str(cluster["citationCluster"])
        print "formatted Citation: " + str(cluster["formattedText"])
