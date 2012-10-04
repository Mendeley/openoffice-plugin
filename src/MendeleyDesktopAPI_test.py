import MendeleyDesktopAPI 
import unittest
import re
import sys

# simplejson is json
try: import simplejson as json
except ImportError: import json

class TestMendeleyDesktopAPI(unittest.TestCase):
    def setUp(self):
        self.api = MendeleyDesktopAPI.MendeleyDesktopAPI("component context (unused)")

        self.testCluster = \
            {
                "citationItems": [{"uris": ["http://local/documents/?uuid=15d6d1e4-a9ff-4258-88b6-a6d6d6bdc0ed"], "id": "ITEM-1", "itemData": {"title": "Title02", "issued": {"date-parts": [["2002"]]}, "author": [{"dropping-particle": "", "non-dropping-particle": "", "parse-names": False, "suffix": "", "given": "Gareth", "family": "Evans"}, {"dropping-particle": "", "given": "Gareth Evans", "family": "Jr", "non-dropping-particle": "", "parse-names": False, "suffix": ""}], "type": "article", "id": "ITEM-1"}}],
                "properties": {"noteIndex": 0},
                "schema": "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"
            }
        self.testUuid = '15d6d1e4-a9ff-4258-88b6-a6d6d6bdc0ed'
        self.testUuid2 = '55ff8735-3f3c-4c9f-87c3-8db322ba3f74'

    class NameValuePair:
        def __init__(self, name, value):
            self.Name = name
            self.Value = value

    def test_execute_multiple_args(self):
        statement = []
        statement.append(self.NameValuePair("functionName", "concatenateStringsTest"))
        statement.append(self.NameValuePair("arg1", "Hello "))
        statement.append(self.NameValuePair("arg2", "World"))
        concatenated = self.api.execute(statement)
        self.assertEqual(concatenated, "Hello World")

    def test_execute(self):
        # set number
        statement = []
        statement.append(self.NameValuePair("functionName", "setNumberTest"))
        statement.append(self.NameValuePair("arg1", "4"))
        response = self.api.execute(statement)
        self.assertEqual(response, "")

        # get number
        statement = []
        statement.append(self.NameValuePair("functionName", "getNumberTest"))
        response = self.api.execute(statement)
        self.assertEqual(response, "4")

    def test__citationClusterFromFieldCode(self):
        # new JSON parsable citaitons should be put in "citaitonCluster" as an
        # object
        newStyleFieldCode = 'ADDIN CSL_CITATION {"testClusterKey": "testClusterValue"}'
        newCitationCluster = self.api._citationClusterFromFieldCode(newStyleFieldCode)
        self.assertEqual(
            newCitationCluster,
            {
                "citationCluster" : 
                {
                    "testClusterKey" : "testClusterValue"
                }
            })
        
        # old non-JSON parsable citations should be put in "fieldCode" as a
        # string
        oldStyleFieldCode = "Mendeley Citation{15d6d1e4-a9ff-4258-88b6-a6d6d6bdc0ed}"
        oldCitationCluster = self.api._citationClusterFromFieldCode(oldStyleFieldCode)
        self.assertEqual(
            oldCitationCluster, {"fieldCode" : oldStyleFieldCode})
        
    def test__fieldCodeFromCitationCluster(self):
        fieldCode = self.api._fieldCodeFromCitationCluster({"testClusterKey": "testClusterValue"})
        self.assertEqual(fieldCode, 'ADDIN CSL_CITATION {"testClusterKey": "testClusterValue"}')

    def test_testGetFieldCodeFromUuid(self):
        fieldCode = self.api.getFieldCodeFromUuid("{" + self.testUuid + "}")
        self.assertEqual(fieldCode, "ADDIN CSL_CITATION " +
            json.dumps(self.testCluster, sort_keys=True))
    
    def test_getUserAccount(self):
        userAccount = self.api.getUserAccount()
        self.assertEqual(userAccount, "testDatabase@test.com@local")

    def test_citation_update_interactive(self):
        updatedCitation = self.api.citation_update_interactive(
            "Mendeley Citation{15d6d1e4-a9ff-4258-88b6-a6d6d6bdc0ed}", "displayed Text")

        self.assertEqual(updatedCitation,
            'ADDIN CSL_CITATION ' + json.dumps(self.testCluster, sort_keys=True))

    def test_citation_undoManualFormat(self):
        citation = self.api.citation_undoManualFormat(
            "Mendeley Edited Citation{" + self.testUuid + "}")
        
        self.assertEqual(citation,
            'ADDIN CSL_CITATION ' +
            json.dumps(self.testCluster, sort_keys=True))

        # invariant on subsequent undos
        self.assertEqual(citation, self.api.citation_undoManualFormat(citation))

    def test_citations_merge(self):
        # merging two identical citations should remain the same
        fieldCodes = []

        fieldCodes.append('ADDIN CSL_CITATION ' + json.dumps(self.testCluster, sort_keys=True))
        fieldCodes.append('ADDIN CSL_CITATION ' + json.dumps(self.testCluster, sort_keys=True))

        mergedFieldCode = self.api.citations_merge(fieldCodes[0], fieldCodes[1])
        self.assertEqual(mergedFieldCode, fieldCodes[0])

        # TODO: test merging different citation clusters
        fieldCodes[1] = 'Mendeley Citation{' + self.testUuid2 + '}'

        mergedFieldCode = self.api.citations_merge(fieldCodes[0], fieldCodes[1])
        self.assertEqual(mergedFieldCode, 'ADDIN CSL_CITATION {"citationItems": [{"id": "ITEM-1", "itemData": {"author": [{"dropping-particle": "", "family": "Evans", "given": "Gareth", "non-dropping-particle": "", "parse-names": false, "suffix": ""}, {"dropping-particle": "", "family": "Jr", "given": "Gareth Evans", "non-dropping-particle": "", "parse-names": false, "suffix": ""}], "id": "ITEM-1", "issued": {"date-parts": [["2002"]]}, "title": "Title02", "type": "article"}, "uris": ["http://local/documents/?uuid=15d6d1e4-a9ff-4258-88b6-a6d6d6bdc0ed"]}, {"id": "ITEM-2", "itemData": {"author": [{"dropping-particle": "", "family": "Smith", "given": "John", "non-dropping-particle": "", "parse-names": false, "suffix": ""}, {"dropping-particle": "", "family": "Jr", "given": "John Smith", "non-dropping-particle": "", "parse-names": false, "suffix": ""}], "id": "ITEM-2", "issued": {"date-parts": [["2001"]]}, "title": "Title01", "type": "article"}, "uris": ["http://local/documents/?uuid=55ff8735-3f3c-4c9f-87c3-8db322ba3f74"]}], "properties": {"noteIndex": 0}, "schema": "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"}')

    def test_wordProcessor_set(self):
        response = self.api.wordProcessor_set("WinWord", 14.0)
        self.assertEqual(response, "")

    def test_formatCitationsAndBibliography(self):
        self.api.resetCitations()
        self.api.setCitationStyle("http://www.zotero.org/styles/apa")
        self.api.addCitationCluster("ADDIN any old text can go here CSL_CITATION { \"citationItems\" : [ { \"id\" : \"ITEM-1\", \"itemData\" : { \"author\" : [ { \"family\" : \"Smith\", \"given\" : \"John\" }, { \"family\" : \"Jr\", \"given\" : \"John Smith\" } ], \"id\" : \"ITEM-1\", \"issued\" : { \"date-parts\" : [ [ \"2001\" ] ] }, \"title\" : \"Title01\", \"type\" : \"article\" }, \"uris\" : [ \"http://local/documents/?uuid=55ff8735-3f3c-4c9f-87c3-8db322ba3f74\" ] }, { \"id\" : \"ITEM-2\", \"itemData\" : { \"author\" : [ { \"family\" : \"Evans\", \"given\" : \"Gareth\" }, { \"family\" : \"Jr\", \"given\" : \"Gareth Evans\" } ], \"id\" : \"ITEM-2\", \"issued\" : { \"date-parts\" : [ [ \"2002\" ] ] }, \"title\" : \"Title02\", \"type\" : \"article\" }, \"uris\" : [ \"http://local/documents/?uuid=15d6d1e4-a9ff-4258-88b6-a6d6d6bdc0ed\" ] } ], \"mendeley\" : { \"previouslyFormattedCitation\" : \"(Evans & Jr, 2002; Smith & Jr, 2001)\" }, \"properties\" : { \"noteIndex\" : 0 }, \"schema\" : \"https://github.com/citation-style-language/schema/raw/master/csl-citation.json\" }")
        self.api.addFormattedCitation("(Evans & Jr, 2002; Smith & Jr, 2001)")
        self.api.addCitationCluster("Mendeley Citation{15d6d1e4-a9ff-4258-88b6-a6d6d6bdc0ed}")
        self.api.addFormattedCitation("test")
        self.api.formatCitationsAndBibliography()

        self.assertEqual(self.citationToJsonString(self.api.getCitationCluster(0)), self.citationToJsonString('ADDIN CSL_CITATION {"mendeley": {"previouslyFormattedCitation": "(Evans & Jr, 2002; Smith & Jr, 2001)"}, "citationItems": [{"uris": ["http://local/documents/?uuid=55ff8735-3f3c-4c9f-87c3-8db322ba3f74"], "id": "ITEM-1", "itemData": {"title": "Title01", "issued": {"date-parts": [["2001"]]}, "author": [{"dropping-particle": "", "given": "John", "family": "Smith", "non-dropping-particle": "", "parse-names": false, "suffix": ""}, {"dropping-particle": "", "given": "John Smith", "family": "Jr", "non-dropping-particle": "", "parse-names": false, "suffix": ""}], "type": "article", "id": "ITEM-1"}}, {"uris": ["http://local/documents/?uuid=15d6d1e4-a9ff-4258-88b6-a6d6d6bdc0ed"], "id": "ITEM-2", "itemData": {"title": "Title02", "issued": {"date-parts": [["2002"]]}, "author": [{"dropping-particle": "", "given": "Gareth", "family": "Evans", "non-dropping-particle": "", "parse-names": false, "suffix": ""}, {"dropping-particle": "", "given": "Gareth Evans", "family": "Jr", "non-dropping-particle": "", "parse-names": false, "suffix": ""}], "type": "article", "id": "ITEM-2"}}], "properties": {"noteIndex": 0}, "schema": "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"}'))
        self.assertEqual(self.api.getFormattedCitation(0), '(Evans & Jr, 2002; Smith & Jr, 2001)')
   
    def test_mendeleyDesktopInfo(self):
        info = self.api.mendeleyDesktopInfo()

        self.assertEqual(self.api.previousSuccess(), "True")
        self.assertTrue(info['processId'])

    def test_pageNotFound(self):
        request = self.api._client.GetRequest("/nonExistant")
        response = self.api._client.request(request)

        self.assertEqual(self.api.previousSuccess(), "False")
        self.assertTrue(self.api.previousErrorMessage() != "")

    def test_previousResultLength(self):
        statement = []
        statement.append(self.NameValuePair("functionName", "mendeleyDesktopInfo"))
        info = str(self.api.execute(statement))

        statement = []
        statement.append(self.NameValuePair("functionName", "previousResultLength"))
        self.assertEqual(len(info), self.api.previousResultLength())

    def citationToJsonString(self,s):
        s = re.sub("^ADDIN CSL_CITATION ","",s)
        return json.dumps(json.loads(s),sort_keys = True)

if __name__ == '__main__':
    unittest.main()
