import MendeleyDesktopAPI 
import unittest
import json

class TestMendeleyDesktopAPI(unittest.TestCase):
    def setUp(self):
        self.api = MendeleyDesktopAPI.MendeleyDesktopAPI("component context (unused)")

        self.testCluster = {"citationItems": [{"uris": ["http://local/documents/?uuid=15d6d1e4-a9ff-4258-88b6-a6d6d6bdc0ed"], "id": "ITEM-1", "itemData": {"title": "Title02", "issued": {"date-parts": [["2002"]]}, "author": [{"given": "Gareth", "family": "Evans"}, {"given": "Gareth Evans", "family": "Jr"}], "note": "<m:note/>", "type": "article", "id": "ITEM-1"}}], "properties": {"noteIndex": 0}, "schema": "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"}
        self.testUuid = '15d6d1e4-a9ff-4258-88b6-a6d6d6bdc0ed'

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
        userAccount = self.api.getUserAccount("")
        self.assertEqual(userAccount, "testDatabase@test.com@local")

    def test_citation_update_interactive(self):
        self.api.setDisplayedText("displayed text")
        updatedCitation = self.api.citation_update_interactive(
            "Mendeley Citation{15d6d1e4-a9ff-4258-88b6-a6d6d6bdc0ed}")

        self.assertEqual(updatedCitation,
            'ADDIN CSL_CITATION ' + json.dumps(self.testCluster, sort_keys=True))

    def test_citation_undoManualFormat(self):
        citation = self.api.citation_undoManualFormat(
            "Mendeley Edited Citation{" + self.testUuid + "}")
        
        # an empty "mendeley" key gets added
        testCluster = self.testCluster
        testCluster["mendeley"] = {}
        self.assertEqual(citation,
            'ADDIN CSL_CITATION ' + json.dumps(testCluster, sort_keys=True))

        # invariant on subsequent undos
        self.assertEqual(citation, self.api.citation_undoManualFormat(citation))

    def test_wordProcessor_set(self):
        response = self.api.wordProcessor_set("WinWord", 14.0)
        self.assertEqual(response, "")

    def test_formatCitationsAndBibliography(self):
        self.api.resetCitations("")
        self.api.setCitationStyle("http://www.zotero.org/styles/apa")
        self.api.addCitationCluster("ADDIN any old text can go here CSL_CITATION { \"citationItems\" : [ { \"id\" : \"ITEM-1\", \"itemData\" : { \"author\" : [ { \"family\" : \"Smith\", \"given\" : \"John\" }, { \"family\" : \"Jr\", \"given\" : \"John Smith\" } ], \"id\" : \"ITEM-1\", \"issued\" : { \"date-parts\" : [ [ \"2001\" ] ] }, \"title\" : \"Title01\", \"type\" : \"article\" }, \"uris\" : [ \"http://local/documents/?uuid=55ff8735-3f3c-4c9f-87c3-8db322ba3f74\" ] }, { \"id\" : \"ITEM-2\", \"itemData\" : { \"author\" : [ { \"family\" : \"Evans\", \"given\" : \"Gareth\" }, { \"family\" : \"Jr\", \"given\" : \"Gareth Evans\" } ], \"id\" : \"ITEM-2\", \"issued\" : { \"date-parts\" : [ [ \"2002\" ] ] }, \"title\" : \"Title02\", \"type\" : \"article\" }, \"uris\" : [ \"http://local/documents/?uuid=15d6d1e4-a9ff-4258-88b6-a6d6d6bdc0ed\" ] } ], \"mendeley\" : { \"previouslyFormattedCitation\" : \"(Evans & Jr, 2002; Smith & Jr, 2001)\" }, \"properties\" : { \"noteIndex\" : 0 }, \"schema\" : \"https://github.com/citation-style-language/schema/raw/master/csl-citation.json\" }")
        self.api.addFormattedCitation("(Evans & Jr, 2002; Smith & Jr, 2001)")
        self.api.addCitationCluster("Mendeley Citation{15d6d1e4-a9ff-4258-88b6-a6d6d6bdc0ed}")
        self.api.addFormattedCitation("test")
        print "formatted citation and bib: " + self.api.formatCitationsAndBibliography("")

        print "Returned citation JSON: " + self.api.getCitationCluster(0)
        print "Returned formatted citation: " + self.api.getFormattedCitation(0)
        print ""
        print "Returned citation JSON: " + self.api.getCitationCluster(1)
        print "Returned formatted citation: " + self.api.getFormattedCitation(1)
        print ""
        print "Returned bibligraphy: " + self.api.getFormattedBibliography("")
    
if __name__ == '__main__':
    unittest.main()