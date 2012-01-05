import MendeleyHttpClient
import MendeleyRPC

import json
import time

client = MendeleyHttpClient.MendeleyHttpClient()
oldClient = MendeleyRPC.MendeleyRPC("unused")

def slowNewApiCalls():
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

    response = client.formattedCitationsAndBibliography_Interactive(
            "http://www.zotero.org/styles/apa", 
            [
                {
                    "formattedText": "",
                    "citationCluster": json.loads("{\"citationItems\": [{\"uris\": [\"http://local/documents/?uuid=ac45152c-4707-4d3c-928d-2cc59aa386fa\"], \"id\": \"ITEM-1\", \"itemData\": {\"title\": \"Overcoming the obstacles of harvesting and searching digital repositories from federated searching toolkits , and embedding them in VLEs Heriot-Watt University Library\", \"author\": [{\"given\": \"Santiago\", \"family\": \"Chumbe\"}, {\"given\": \"Roddy\", \"family\": \"Macleod\"}, {\"given\": \"Phil\", \"family\": \"Barker\"}, {\"given\": \"Malcolm\", \"family\": \"Moffat\"}, {\"given\": \"Roger\", \"family\":\"Rist\"}], \"note\": \"<m:note/>\", \"container-title\": \"Language\", \"type\": \"article-journal\", \"id\": \"ITEM-1\"}}], \"properties\": {\"noteIndex\": 0}, \"schema\": \"https://github.com/citation-style-language/schema/raw/master/csl-citation.json\"}")
                }
            ]
            )
    for cluster in response.citationClusters:
        print "cluster: " + json.dumps(cluster["citationCluster"])
        print "formatted Citation: " + json.dumps(cluster["formattedText"])

def quickNewApiCalls():
    response = client.getUserAccount()
    print "user account = " + json.dumps(response.__dict__)
    print "request time = " + str(client.lastRequestTime) + "ms"

    response = client.getUserAccount()
    print "user account = " + json.dumps(response.__dict__)
    print "request time = " + str(client.lastRequestTime) + "ms"

    response = client.getUserAccount()
    print "user account = " + json.dumps(response.__dict__)
    print "request time = " + str(client.lastRequestTime) + "ms"


class Value:
    def __init__(self, value):
        self.Value = value

def oldApiCall(functionName, argument):
    startTime = time.time()
    responseLength = oldClient.execute({0: Value(functionName + unichr(13) + argument)})
    response = oldClient.execute({0: Value("getStringResult")})
    print "old Api request: " + functionName + " took " + str(1000 * (time.time() - startTime)) + "ms"
    return response


def makeOldApiCalls():
    oldApiResponse = oldApiCall("getUserAccount", "")
    print "olduser account = " + oldApiResponse

    oldApiResponse = oldApiCall("getUserAccount", "")
    print "olduser account = " + oldApiResponse

    oldApiResponse = oldApiCall("getUserAccount", "")
    print "olduser account = " + oldApiResponse

quickNewApiCalls()
makeOldApiCalls()
quickNewApiCalls()
makeOldApiCalls()
slowNewApiCalls()

