import MendeleyHttpClient
import time
import unittest
# simplejson is json 
try: import simplejson as json
except ImportError: import json

class TestMendeleyHttpClient(unittest.TestCase):

    def setUp(self):
        self.client = MendeleyHttpClient.MendeleyHttpClient()
        # todo: tidy up first is "formattedText" and "citationCluster", second just "citationCluster"
        self.testClusters = [
            '{"formattedText": "(Evans & Jr, 2002; Smith & Jr, 2001)", "citationCluster": {"mendeley": {"formattedCitation": "(Evans & Jr, 2002; Smith & Jr, 2001)", "plainTextFormattedCitation": "(Evans & Jr, 2002; Smith & Jr, 2001)", "previouslyFormattedCitation": "(Evans & Jr, 2002; Smith & Jr, 2001)"}, "citationItems": [{"uris": ["http://local/documents/?uuid=55ff8735-3f3c-4c9f-87c3-8db322ba3f74"], "id": "ITEM-1", "itemData": {"issued": {"date-parts": [["2001"]]}, "title": "Title01", "type": "article", "id": "ITEM-1", "author": [{"given": "John", "dropping-particle": "", "suffix": "", "family": "Smith", "parse-names": false, "non-dropping-particle": ""}, {"given": "John Smith", "dropping-particle": "", "suffix": "", "family": "Jr", "parse-names": false, "non-dropping-particle": ""}]}}, {"uris": ["http://local/documents/?uuid=15d6d1e4-a9ff-4258-88b6-a6d6d6bdc0ed"], "id": "ITEM-2", "itemData": {"issued": {"date-parts": [["2002"]]}, "title": "Title02", "type": "article", "id": "ITEM-2", "author": [{"given": "Gareth", "dropping-particle": "", "suffix": "", "family": "Evans", "parse-names": false, "non-dropping-particle": ""}, {"given": "Gareth Evans", "dropping-particle": "", "suffix": "", "family": "Jr", "parse-names": false, "non-dropping-particle": ""}]}}], "properties": {"noteIndex": 0}, "schema": "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"}}',
            '{"mendeley": {"formattedCitation": "(Evans & Jr, 2002; Smith & Jr, 2001)", "plainTextFormattedCitation": "(Evans & Jr, 2002; Smith & Jr, 2001)", "previouslyFormattedCitation": "(Evans & Jr, 2002; Smith & Jr, 2001)"}, "citationItems": [{"uris": ["http://local/documents/?uuid=55ff8735-3f3c-4c9f-87c3-8db322ba3f74"], "id": "ITEM-1", "itemData": {"issued": {"date-parts": [["2001"]]}, "title": "Title01", "type": "article", "id": "ITEM-1", "author": [{"given": "John", "dropping-particle": "", "suffix": "", "family": "Smith", "parse-names": false, "non-dropping-particle": ""}, {"given": "John Smith", "dropping-particle": "", "suffix": "", "family": "Jr", "parse-names": false, "non-dropping-particle": ""}]}}, {"uris": ["http://local/documents/?uuid=15d6d1e4-a9ff-4258-88b6-a6d6d6bdc0ed"], "id": "ITEM-2", "itemData": {"issued": {"date-parts": [["2002"]]}, "title": "Title02", "type": "article", "id": "ITEM-2", "author": [{"given": "Gareth", "dropping-particle": "", "suffix": "", "family": "Evans", "parse-names": false, "non-dropping-particle": ""}, {"given": "Gareth Evans", "dropping-particle": "", "suffix": "", "family": "Jr", "parse-names": false, "non-dropping-particle": ""}]}}], "properties": {"noteIndex": 0}, "schema": "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"}'
            ]

    def test_simpleNewApiCall(self):
        response = self.client.userAccount()
        self.assertEqual(response.body.account, "testDatabase@test.com@local")

    def test_formatCitationsAndBibliography(self):
        # this call should complete without requiring user interaction

        response1 = self.client.formattedCitationsAndBibliography_Interactive(
                "http://www.zotero.org/styles/apa",
                [
                    json.loads(self.testClusters[0]),
                    json.loads(self.testClusters[0])
                ]
                )

        self.assertEqual(response1.status, 200)

        # should get the same cluster1 back
        self.assertEqual(
                json.dumps(response1.body.citationClusters[0], sort_keys=True),
                self.stringToSortedJson(self.testClusters[0]))

        # should get the same cluster2 back
        self.assertEqual(
                json.dumps(response1.body.citationClusters[1], sort_keys=True),
                self.stringToSortedJson(self.testClusters[0]))

        response2 = self.client.formattedCitationsAndBibliography_Interactive(
                "http://www.zotero.org/styles/apa", 
                [
                    {
                        "citationCluster":json.loads(self.testClusters[1])
                    }
                ]
                )

        self.assertEqual(
                json.dumps(response1.body.citationClusters[0]["citationCluster"], sort_keys=True),
                self.stringToSortedJson(self.testClusters[1]))
        
        self.assertEqual(
                response1.body.citationClusters[0]["formattedText"],
                "(Evans & Jr, 2002; Smith & Jr, 2001)")

        # for now the bibliography is written to a temp file, this may change in future
        bibliography = open(response1.body.bibliography).read()
        
        expected = """{\\rtf
\\par\\sl288\\slmult1\\sb0\\sa140\\li480\\fi-480 Evans, G., & Jr, G. E. (2002). Title02.
\\par\\sl288\\slmult1\\sb0\\sa140\\li480\\fi-480 Smith, J., & Jr, J. S. (2001). Title01.

}"""
        self.assertEqual(
                bibliography,
                expected
                )

    def test_styleName_getFromUrl(self):
        response = self.client.styleName_getFromUrl(
                {"citationStyleUrl": "http://www.zotero.org/styles/apa"})
        self.assertEqual(response.status, 200)
        self.assertEqual(response.body.citationStyleName, "American Psychological Association 6th edition")

    def test_citationStyles_default(self):
        response = self.client.citationStyles_default()
        self.assertEqual(response.status, 200)
        self.assertEqual(len(response.body.citationStyles), 10)
        
        self.assertTrue(len(response.body.citationStyles[0]["title"]) > 0)
        self.assertTrue(len(response.body.citationStyles[0]["url"]) > 0)

    def test_citations_merge(self):
        response = self.client.citations_merge(
            {"citationClusters":
                [
                    {"citationCluster": json.loads(self.testClusters[0])},
                    {"citationCluster": json.loads(self.testClusters[1])}
                ]
            })
        self.assertEqual(response.status, 200)

        self.assertEqual(json.dumps(response.body.citationCluster,sort_keys=True),
            '{"citationItems": [{"id": "ITEM-1", "itemData": {"author": [{"dropping-particle": "", "family": "Smith", "given": "John", "non-dropping-particle": "", "parse-names": false, "suffix": ""}, {"dropping-particle": "", "family": "Jr", "given": "John Smith", "non-dropping-particle": "", "parse-names": false, "suffix": ""}], "id": "ITEM-1", "issued": {"date-parts": [["2001"]]}, "title": "Title01", "type": "article"}, "uris": ["http://local/documents/?uuid=55ff8735-3f3c-4c9f-87c3-8db322ba3f74"]}, {"id": "ITEM-2", "itemData": {"author": [{"dropping-particle": "", "family": "Evans", "given": "Gareth", "non-dropping-particle": "", "parse-names": false, "suffix": ""}, {"dropping-particle": "", "family": "Jr", "given": "Gareth Evans", "non-dropping-particle": "", "parse-names": false, "suffix": ""}], "id": "ITEM-2", "issued": {"date-parts": [["2002"]]}, "title": "Title02", "type": "article"}, "uris": ["http://local/documents/?uuid=15d6d1e4-a9ff-4258-88b6-a6d6d6bdc0ed"]}], "properties": {"noteIndex": 0}, "schema": "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"}')

    def test_citation_undoManualFormat(self):
        citation = json.loads(self.testClusters[0])["citationCluster"]
        citation["mendeley"]["manualFormatting"] = "Test manual format"
        self.assertTrue("manualFormatting" in citation["mendeley"])
        response = self.client.citation_undoManualFormat(
            {"citationCluster" : citation } )
        self.assertEqual(response.status, 200)
        self.assertFalse(
            "mendeley" in response.body.citationCluster and
            "manualFormatting" in response.body.citationCluster["mendeley"])
        
    def test_wordProcessor_set(self):
        response = self.client.wordProcessor_set(
            {
                "wordProcessor": "test processor",
                "version": 999
            })
        self.assertEqual(response.status, 200)
        
        # incomplete request should generate 400 response
        response = self.client.wordProcessor_set(
            {
                "wrong key" : "no!"
            })
        self.assertEqual(response.status, 400)

    def stringToSortedJson(self,s):
        return json.dumps(json.loads(s), sort_keys = True)

if __name__ == '__main__':
    unittest.main()
