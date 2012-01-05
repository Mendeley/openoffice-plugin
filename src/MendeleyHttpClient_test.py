import MendeleyHttpClient
import MendeleyRPC

import json
import time

import unittest

class OldApiClient():
    def __init__(self):
        self.oldClient = MendeleyRPC.MendeleyRPC("unused")

    class Value:
        def __init__(self, value):
            self.Value = value

    def request(self, functionName, argument):
        startTime = time.time()
        responseLength = self.oldClient.execute({0: self.Value(functionName + unichr(13) + argument)})
        response = self.oldClient.execute({0: self.Value("getStringResult")})
        print "old Api request: " + functionName + " took " + str(1000 * (time.time() - startTime)) + "ms"
        return response

class TestMendeleyHttpClient(unittest.TestCase):

    def setUp(self):
        self.client = MendeleyHttpClient.MendeleyHttpClient()
        self.oldClient = OldApiClient()

    def test_simpleOldApiCall(self):
        userAccount = self.oldClient.request("getUserAccount", "")
        self.assertEqual(userAccount, "testDatabase@test.com@local")

    def test_simpleNewApiCall(self):
        response = self.client.getUserAccount()
        self.assertEqual(response.account, "testDatabase@test.com@local")

    def test_slowNewApiCalls(self):
        # this call should complete without requiring user interaction

        jsonCluster1 = '{"formattedText": "(Evans & Jr, 2002; Smith & Jr, 2001)", "citationCluster": {"mendeley": {"previouslyFormattedCitation": "(Evans & Jr, 2002; Smith & Jr, 2001)"}, "citationItems": [{"uris": ["http://local/documents/?uuid=55ff8735-3f3c-4c9f-87c3-8db322ba3f74"], "id": "ITEM-1", "itemData": {"title": "Title01", "issued": {"date-parts": [["2001"]]}, "author": [{"given": "John", "family": "Smith"}, {"given": "John Smith", "family": "Jr"}], "note": "<m:note/>", "type": "article", "id": "ITEM-1"}}, {"uris": ["http://local/documents/?uuid=15d6d1e4-a9ff-4258-88b6-a6d6d6bdc0ed"], "id": "ITEM-2", "itemData": {"title": "Title02", "issued": {"date-parts": [["2002"]]}, "author": [{"given": "Gareth", "family": "Evans"}, {"given": "Gareth Evans", "family": "Jr"}], "note": "<m:note/>", "type": "article", "id": "ITEM-2"}}], "properties": {"noteIndex": 0}, "schema": "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"}}'
        jsonCluster2 = '{"mendeley": {"previouslyFormattedCitation": "(Evans & Jr, 2002; Smith & Jr, 2001)"}, "citationItems": [{"uris": ["http://local/documents/?uuid=55ff8735-3f3c-4c9f-87c3-8db322ba3f74"], "id": "ITEM-1", "itemData": {"title": "Title01", "issued": {"date-parts": [["2001"]]}, "author": [{"given": "John", "family": "Smith"}, {"given": "John Smith", "family": "Jr"}], "note": "<m:note/>", "type": "article", "id": "ITEM-1"}}, {"uris": ["http://local/documents/?uuid=15d6d1e4-a9ff-4258-88b6-a6d6d6bdc0ed"], "id": "ITEM-2", "itemData": {"title": "Title02", "issued": {"date-parts": [["2002"]]}, "author": [{"given": "Gareth", "family": "Evans"}, {"given": "Gareth Evans", "family": "Jr"}], "note": "<m:note/>", "type": "article", "id": "ITEM-2"}}], "properties": {"noteIndex": 0}, "schema": "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"}'

        response1 = self.client.formattedCitationsAndBibliography_Interactive(
                "http://www.zotero.org/styles/apa",
                [
                    json.loads(jsonCluster1),
                    json.loads(jsonCluster1)
                ]
                )

        # should get the same cluster1 back
        self.assertEqual(
                json.dumps(response1.citationClusters[0]),
                jsonCluster1)

        # should get the same cluster2 back
        self.assertEqual(
                json.dumps(response1.citationClusters[1]),
                jsonCluster1)

        response2 = self.client.formattedCitationsAndBibliography_Interactive(
                "http://www.zotero.org/styles/apa", 
                [
                    {
                        "citationCluster":json.loads(jsonCluster2)
                    }
                ]
                )

        self.assertEqual(
                json.dumps(response1.citationClusters[0]["citationCluster"]),
                jsonCluster2)
        
        self.assertEqual(
                response1.citationClusters[0]["formattedText"],
                "(Evans & Jr, 2002; Smith & Jr, 2001)")

        # for now the bibliography is written to a temp file, this may change in future
        bibliography = open(response1.bibliography).read()
        self.assertEqual(
                bibliography,
                '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">\n<html xmlns="http://www.w3.org/1999/xhtml"><head>\n<meta http-equiv="Content-Type" content="text/html;charset=utf-8" /><title></title>\n</head>\n<body>\n&nbsp;<p>\n<p style=\'margin-left:24pt;text-indent:-24.0pt\'>Evans, G., &#38; Jr, G. E. (2002). Title02.</p><p style=\'margin-left:24pt;text-indent:-24.0pt\'>Smith, J., &#38; Jr, J. S. (2001). Title01.</p>\n</p></body></html>\n'
                )

if __name__ == '__main__':
    unittest.main()
