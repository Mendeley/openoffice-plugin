import MendeleyDesktopAPI 

import unittest

class TestMendeleyDesktopAPI(unittest.TestCase):

    def setUp(self):
        self.api = MendeleyDesktopAPI.MendeleyDesktopAPI("component context (unused)")

    def test__fieldCodeFromCitationCluster(self):
        fieldCode = self.api._fieldCodeFromCitationCluster({"testClusterKey": "testClusterValue"})
        self.assertEqual(fieldCode, 'ADDIN CSL_CITATION {"testClusterKey": "testClusterValue"}')

    def test_testGetFieldCodeFromUuid(self):
        fieldCode = self.api.getFieldCodeFromUuid("{15d6d1e4-a9ff-4258-88b6-a6d6d6bdc0ed}")
        print fieldCode

    def test_
    
    def test_getUserAccount(self):
        userAccount = self.api.getUserAccount("")
        self.assertEqual(userAccount, "testDatabase@test.com@local")

    def test_formatCitationsAndBibliography(self):
        print "Format some citations and a bibliography"
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
